"""线程池拥有权 Mixin

PDFConverter 和 OfficeGenerator 存在相同的 executor / _owns_executor / cleanup 模式。
提取到此 mixin 以消除重复代码。
"""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor

from astrbot.api import logger


class ExecutorOwnerMixin:
    """管理可选的外部/内部线程池的生命周期。"""

    _executor: ThreadPoolExecutor | None = None
    _owns_executor: bool = False
    _executor_label: str = ""

    def _init_executor(
        self,
        executor: ThreadPoolExecutor | None,
        *,
        max_workers: int = 2,
        label: str = "",
    ) -> None:
        self._owns_executor = executor is None
        self._executor = (
            executor
            if executor is not None
            else ThreadPoolExecutor(max_workers=max_workers)
        )
        self._executor_label = label

    def _shutdown_executor(self) -> None:
        executor = self._executor
        if self._owns_executor and executor is not None:
            label = self._executor_label
            executor.shutdown(wait=False)
            self._executor = None
            self._owns_executor = False
            self._executor_label = ""
            if label:
                logger.debug(f"[{label}] 线程池已关闭")

    def _require_executor(self) -> ThreadPoolExecutor:
        executor = self._executor
        if executor is None:
            raise RuntimeError(
                "ExecutorOwnerMixin requires _init_executor() before use"
            )
        return executor
