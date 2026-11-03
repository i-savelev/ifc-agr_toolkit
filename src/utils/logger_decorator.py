"""Декораторы, использующие проектный логгер."""

from functools import wraps
from typing import Any, Callable

import pandas as pd

from .logger import Logger as log


def log_dataframe(func: Callable[..., pd.DataFrame]) -> Callable[..., pd.DataFrame]:
    """Логирует DataFrame, возвращаемый функцией.

    :param func: Функция, результатом которой является DataFrame.
    :returns: Обёрнутая функция с тем же результатом.
    """

    @wraps(func)
    def wrapper(*args: Any, **kwargs: Any) -> pd.DataFrame:
        result = func(*args, **kwargs)
        if isinstance(result, pd.DataFrame):
            log.raw_dataframe(result, caption=f"Результат выполнения '{func.__name__}'")
        return result

    return wrapper
