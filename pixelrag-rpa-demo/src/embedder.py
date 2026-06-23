"""图像向量化模块

支持两种模式：
- real: 使用 Qwen3-VL-Embedding-2B 模型（需要 torch + transformers）
- mock: 使用图像直方图 + 感知哈希作为特征向量（无需额外依赖）
"""

import hashlib
import numpy as np
from PIL import Image


class ImageEmbedder:
    """图像向量化器"""

    def __init__(self, mode: str = "auto", model_path: str | None = None):
        """
        Args:
            mode: "auto" | "real" | "mock"
            model_path: 模型路径（real 模式）
        """
        self.mode = mode
        self.model_path = model_path
        self._model = None
        self._processor = None
        self._device = None
        self._dimension = None
        self._init()

    def _init(self):
        if self.mode == "real":
            self._init_real()
        elif self.mode == "mock":
            self._init_mock()
        else:  # auto
            try:
                self._init_real()
                self.mode = "real"
            except Exception:
                self._init_mock()
                self.mode = "mock"

    def _init_real(self):
        import torch
        from transformers import AutoProcessor, Qwen3VLForConditionalGeneration

        device = "cuda" if torch.cuda.is_available() else "cpu"
        dtype = torch.float32 if device == "cpu" else torch.bfloat16
        model_name = self.model_path or "Qwen/Qwen3-VL-Embedding-2B"

        processor = AutoProcessor.from_pretrained(model_name, trust_remote_code=True)
        model = Qwen3VLForConditionalGeneration.from_pretrained(
            model_name, trust_remote_code=True, dtype=dtype
        )
        model = model.to(device).eval()

        self._model = model
        self._processor = processor
        self._device = device
        self._dimension = model.config.hidden_size

    def _init_mock(self):
        self._dimension = 256

    @property
    def dimension(self) -> int:
        return self._dimension

    def embed(self, image_path: str) -> np.ndarray:
        """将单张图片转为向量"""
        img = Image.open(image_path).convert("RGB")
        if self.mode == "real":
            return self._embed_real(img)
        else:
            return self._embed_mock(img)

    def embed_batch(self, image_paths: list[str]) -> np.ndarray:
        """批量向量化"""
        vectors = []
        for p in image_paths:
            vectors.append(self.embed(p))
        return np.stack(vectors)

    def _embed_real(self, img: Image.Image) -> np.ndarray:
        import torch

        messages = [
            {"role": "system", "content": [{"type": "text", "text": "Represent the user's input."}]},
            {"role": "user", "content": [{"type": "image", "image": img}]},
        ]
        text = self._processor.apply_chat_template(
            messages, tokenize=False, add_generation_prompt=True
        )
        inputs = self._processor(
            text=[text], images=[img], return_tensors="pt", padding=True
        )
        inputs = {k: v.to(self._device) if hasattr(v, "to") else v for k, v in inputs.items()}

        with torch.no_grad():
            outputs = self._model.model(**inputs)

        last_hidden = outputs.last_hidden_state
        attention_mask = inputs["attention_mask"]
        last_token_indices = attention_mask.sum(dim=1) - 1
        pooled = last_hidden[
            torch.arange(last_hidden.size(0), device=last_hidden.device),
            last_token_indices,
        ]
        pooled = torch.nn.functional.normalize(pooled, p=2, dim=-1)
        return pooled.cpu().float().numpy()[0].astype(np.float32)

    def _embed_mock(self, img: Image.Image) -> np.ndarray:
        """模拟向量化：颜色直方图 + 感知哈希 + 图像统计特征

        生成 256 维向量，L2 归一化。
        虽然不如真实模型准确，但足够演示检索流程。
        """
        # 1. 颜色直方图（192 维 = 64*3）
        img_small = img.resize((64, 64))
        arr = np.array(img_small, dtype=np.float32)

        r_hist = np.histogram(arr[:, :, 0], bins=64, range=(0, 256))[0]
        g_hist = np.histogram(arr[:, :, 1], bins=64, range=(0, 256))[0]
        b_hist = np.histogram(arr[:, :, 2], bins=64, range=(0, 256))[0]
        hist = np.concatenate([r_hist, g_hist, b_hist]).astype(np.float32)

        # 2. 感知哈希特征（49 维 = 7x7）
        hash_img = img.convert("L").resize((8, 8), Image.LANCZOS)
        hash_arr = np.array(hash_img, dtype=np.float32)
        mean_val = hash_arr.mean()
        hash_bits = (hash_arr > mean_val).astype(np.float32).flatten()
        hash_bits = hash_bits[:49]

        # 3. 统计特征（15 维）
        stats = np.array([
            arr[:, :, 0].mean(), arr[:, :, 0].std(),
            arr[:, :, 1].mean(), arr[:, :, 1].std(),
            arr[:, :, 2].mean(), arr[:, :, 2].std(),
            float(arr.shape[0]), float(arr.shape[1]),
            float(arr[:, :, 0].max()), float(arr[:, :, 1].max()),
            float(arr[:, :, 2].max()), float(arr[:, :, 0].min()),
            float(arr[:, :, 1].min()), float(arr[:, :, 2].min()),
            float(np.mean(np.abs(arr[:, :, 0] - arr[:, :, 1]))),
        ], dtype=np.float32)

        # 拼接并归一化
        vec = np.concatenate([hist, hash_bits, stats])
        vec = vec[: self._dimension]
        if len(vec) < self._dimension:
            vec = np.pad(vec, (0, self._dimension - len(vec)))

        norm = np.linalg.norm(vec)
        if norm > 1e-12:
            vec = vec / norm
        return vec.astype(np.float32)
