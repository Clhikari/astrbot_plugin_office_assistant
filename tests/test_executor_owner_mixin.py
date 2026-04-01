from concurrent.futures import ThreadPoolExecutor

import pytest

from astrbot_plugin_office_assistant._executor_mixin import ExecutorOwnerMixin


class _DummyExecutorOwner(ExecutorOwnerMixin):
    pass


def test_shutdown_executor_clears_owned_executor_reference():
    owner = _DummyExecutorOwner()
    owner._init_executor(None, label="测试线程池")

    owner._shutdown_executor()

    assert owner._executor is None
    assert owner._owns_executor is False
    assert owner._executor_label == ""
    with pytest.raises(RuntimeError, match="_init_executor"):
        owner._require_executor()


def test_shutdown_executor_keeps_external_executor_reference():
    owner = _DummyExecutorOwner()
    external_executor = ThreadPoolExecutor(max_workers=1)
    try:
        owner._init_executor(external_executor, label="外部线程池")

        owner._shutdown_executor()

        assert owner._executor is external_executor
        assert owner._owns_executor is False
        assert owner._executor_label == "外部线程池"
        assert owner._require_executor() is external_executor
    finally:
        external_executor.shutdown(wait=False)
