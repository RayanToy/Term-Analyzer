"""
Патч для совместимости pymorphy2 (внутри natasha) с новыми версиями Python.
"""

import sys


def apply():
    """Применяет патч если pkg_resources недоступен или неполный."""

    # Пробуем настоящий pkg_resources
    try:
        import pkg_resources
        # Проверяем что WorkingSet есть
        _ = pkg_resources.WorkingSet
        return  # всё хорошо — ничего не делаем
    except (ImportError, AttributeError):
        pass

    # Строим полноценную заглушку
    import types
    from importlib.metadata import entry_points

    def _get_entry_points(group, name=None):
        try:
            eps = entry_points()
            if hasattr(eps, 'select'):
                result = list(eps.select(group=group))
            else:
                result = list(eps.get(group, []))
            if name is not None:
                result = [ep for ep in result if ep.name == name]
            return result
        except Exception:
            return []

    class _Distribution:
        """Минимальная заглушка Distribution."""
        def __init__(self, ep):
            self._ep = ep

        def get_entry_map(self, group=None):
            return {}

    class _WorkingSet:
        """
        Заглушка WorkingSet — pymorphy2 использует её для поиска
        зарегистрированных словарей через entry_points.
        """
        def __init__(self, entries=None):
            pass

        def iter_entry_points(self, group, name=None):
            return iter(_get_entry_points(group, name))

        def __iter__(self):
            return iter([])

    mod = types.ModuleType('pkg_resources')
    mod.WorkingSet = _WorkingSet
    mod.iter_entry_points = _get_entry_points
    mod.DistributionNotFound = Exception
    mod.VersionConflict = Exception
    mod.require = lambda *a, **kw: None
    sys.modules['pkg_resources'] = mod
    print("[патч] pkg_resources заглушка установлена (WorkingSet включён)")