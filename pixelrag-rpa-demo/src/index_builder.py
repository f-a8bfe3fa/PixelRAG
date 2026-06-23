"""索引构建模块：将案例库中的异常截图向量化并构建 FAISS 索引"""

import json
import os
import time
from pathlib import Path

import faiss
import numpy as np

from .embedder import ImageEmbedder


class CaseIndex:
    """案例索引管理器"""

    def __init__(self, cases_dir: str, index_dir: str, embed_mode: str = "auto"):
        self.cases_dir = Path(cases_dir)
        self.index_dir = Path(index_dir)
        self.embed_mode = embed_mode
        self.embedder = None
        self.index = None
        self.case_ids = []
        self._meta = {}

    def build(self):
        """扫描案例库，向量化，构建 FAISS 索引"""
        print(f"扫描案例库: {self.cases_dir}")
        case_dirs = sorted([
            d for d in self.cases_dir.iterdir()
            if d.is_dir() and (d / "before.png").exists()
        ])

        if not case_dirs:
            print("未找到案例！")
            return

        print(f"找到 {len(case_dirs)} 个案例")

        self.embedder = ImageEmbedder(mode=self.embed_mode)
        print(f"向量化模式: {self.embedder.mode} (维度={self.embedder.dimension})")

        before_images = []
        case_ids = []
        metas = []

        for case_dir in case_dirs:
            case_id = case_dir.name
            before_path = case_dir / "before.png"
            meta_path = case_dir / "meta.json"

            case_ids.append(case_id)
            before_images.append(str(before_path))

            if meta_path.exists():
                with open(meta_path, "r", encoding="utf-8") as f:
                    metas.append(json.load(f))
            else:
                metas.append({"case_id": case_id})

        # 向量化
        print("正在向量化图片...")
        t0 = time.time()
        embeddings = self.embedder.embed_batch(before_images)
        print(f"向量化完成，耗时 {time.time() - t0:.2f}s")

        # 构建 FAISS 索引
        dim = embeddings.shape[1]
        index = faiss.IndexFlatIP(dim)
        index.add(embeddings.astype(np.float32))

        # 保存
        self.index_dir.mkdir(parents=True, exist_ok=True)

        index_path = self.index_dir / "index.faiss"
        faiss.write_index(index, str(index_path))
        print(f"索引已保存: {index_path} ({index.ntotal} 条向量)")

        # 保存元数据
        meta_path = self.index_dir / "metadata.npz"
        case_ids_arr = np.array(case_ids)
        np.savez(str(meta_path), case_ids=case_ids_arr)
        print(f"元数据已保存: {meta_path}")

        # 保存摘要
        summary = {
            "total_cases": len(case_ids),
            "dimension": dim,
            "embed_mode": self.embedder.mode,
            "index_type": "FlatIP",
            "case_ids": case_ids,
        }
        summary_path = self.index_dir / "summary.json"
        with open(summary_path, "w", encoding="utf-8") as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)
        print(f"摘要已保存: {summary_path}")

    def load(self):
        """加载已构建的索引"""
        index_path = self.index_dir / "index.faiss"
        meta_path = self.index_dir / "metadata.npz"
        summary_path = self.index_dir / "summary.json"

        if not index_path.exists():
            raise FileNotFoundError(f"索引不存在: {index_path}，请先运行 build()")

        self.index = faiss.read_index(str(index_path))
        meta = np.load(str(meta_path), allow_pickle=True)
        self.case_ids = list(meta["case_ids"])

        if summary_path.exists():
            with open(summary_path, "r", encoding="utf-8") as f:
                self._meta = json.load(f)

        self.embedder = ImageEmbedder(mode=self._meta.get("embed_mode", "auto"))
        print(f"索引已加载: {self.index.ntotal} 条向量, 维度={self.index.d}")

    def search(self, query_image_path: str, top_k: int = 5) -> list[dict]:
        """搜索相似案例

        Returns:
            [{"case_id": "...", "score": 0.95, "rank": 1}, ...]
        """
        if self.index is None:
            self.load()

        query_vec = self.embedder.embed(query_image_path).reshape(1, -1)
        query_vec = query_vec.astype(np.float32)

        # L2 归一化（FlatIP = cosine similarity for normalized vectors）
        norms = np.linalg.norm(query_vec, axis=1, keepdims=True)
        query_vec = query_vec / np.maximum(norms, 1e-12)

        distances, indices = self.index.search(query_vec, top_k)

        results = []
        for i in range(min(top_k, len(indices[0]))):
            idx = indices[0][i]
            if idx < 0 or idx >= len(self.case_ids):
                continue
            results.append({
                "rank": i + 1,
                "case_id": self.case_ids[idx],
                "score": float(distances[0][i]),
            })

        return results

    def get_case_dir(self, case_id: str) -> Path:
        return self.cases_dir / case_id
