"""
Централизованный реестр входных Excel-отчётов Wildberries и карта соответствия колонок.

Каждый отчёт описан:
- обязательностью (required);
- группой в интерфейсе (group): обязательные / дополнительные / ручные справочники;
- связью с аналитическими блоками (blocks);
- подсказкой, где скачать отчёт в кабинете (where);
- пояснением, что именно будет недоступно без этого отчёта (unavailable_if_missing);
- признаком мультизагрузки (supports_multiple) — например, для недельных финансовых отчётов.

Тексты рассчитаны на обычного пользователя без технической подготовки.
"""

# Группы для интерфейса загрузки.
GROUP_REQUIRED = "required"     # Обязательные данные
GROUP_OPTIONAL = "optional"     # Дополнительные данные (выгрузки из кабинета WB)
GROUP_MANUAL = "manual"         # Ручные справочники (заполняются пользователем)

REPORTS_REGISTRY = {
    "finance_weekly": {
        "title": "Еженедельный финансовый отчёт (детализация)",
        "required": True,
        "group": GROUP_REQUIRED,
        "supports_multiple": True,
        "blocks": ["sales", "finance", "unit_economics_sku", "unit_economics_cabinet"],
        "where": "Финансовые отчёты → Еженедельные отчёты → Скачать детализацию (файл за каждую неделю)",
        "unavailable_if_missing": (
            "Это базовый отчёт. Без него не рассчитываются продажи, финансы и юнит-экономика."
        ),
    },
    "sales_report": {
        "title": "Отчёт «Продажи»",
        "required": False,
        "group": GROUP_OPTIONAL,
        "supports_multiple": False,
        "blocks": ["sales", "trends"],
        "where": "Аналитика → Отчёты → Продажи",
        "unavailable_if_missing": "Без него не будет расширенной аналитики трендов продаж.",
    },
    "stocks_report": {
        "title": "Отчёт по остаткам на складах",
        "required": False,
        "group": GROUP_OPTIONAL,
        "supports_multiple": False,
        "blocks": ["stocks", "warehouses"],
        "where": "Аналитика → Отчёты → Остатки на складах",
        "unavailable_if_missing": "Без него не рассчитываются остатки и риски дефицита товара.",
    },
    "turnover_report": {
        "title": "Отчёт «Оборачиваемость»",
        "required": False,
        "group": GROUP_OPTIONAL,
        "supports_multiple": False,
        "blocks": ["stocks", "risks"],
        "where": "Аналитика → Отчёты → Оборачиваемость",
        "unavailable_if_missing": "Без него не будет анализа оборачиваемости и части блока рисков.",
    },
    "nomenclature_report": {
        "title": "Перечень номенклатур (карточки товаров)",
        "required": False,
        "group": GROUP_OPTIONAL,
        "supports_multiple": False,
        "blocks": ["sales", "unit_economics_sku"],
        "where": "Товары → Карточки товаров → Экспорт",
        "unavailable_if_missing": "Без него сложнее сопоставлять артикулы с названиями товаров.",
    },
    "storage_report": {
        "title": "Отчёт «Платное хранение»",
        "required": False,
        "group": GROUP_OPTIONAL,
        "supports_multiple": False,
        "blocks": ["finance", "unit_economics_sku"],
        "where": "Аналитика → Отчёты → Платное хранение",
        "unavailable_if_missing": "Без него расходы на платное хранение не будут учтены отдельно.",
    },
    "ads_report": {
        "title": "История затрат на продвижение (реклама WB)",
        "required": False,
        "group": GROUP_OPTIONAL,
        "supports_multiple": False,
        "blocks": ["promotion", "unit_economics_sku"],
        "where": "Продвижение → Реклама → Финансы → История затрат",
        "unavailable_if_missing": "Без него рекламные расходы из кабинета не попадут в юнит-экономику.",
    },
    "cost_price_manual": {
        "title": "Себестоимость SKU (ручная таблица)",
        "required": False,
        "group": GROUP_MANUAL,
        "supports_multiple": False,
        "blocks": ["unit_economics_sku", "unit_economics_cabinet"],
        "where": "Заполняется вручную по шаблону «Себестоимость SKU»",
        "unavailable_if_missing": (
            "Без неё не рассчитывается чистая прибыль по SKU — доступна только валовая выручка."
        ),
    },
    "extra_ads_manual": {
        "title": "Дополнительные рекламные расходы (ручная таблица)",
        "required": False,
        "group": GROUP_MANUAL,
        "supports_multiple": False,
        "blocks": ["promotion", "unit_economics_sku"],
        "where": "Заполняется вручную по шаблону «Дополнительные рекламные расходы»",
        "unavailable_if_missing": "Без него не учитываются рекламные расходы вне кабинета Wildberries.",
    },
    "classifier_manual": {
        "title": "Классификатор SKU / Категория / Бренд (ручная таблица)",
        "required": False,
        "group": GROUP_MANUAL,
        "supports_multiple": False,
        "blocks": ["sales", "unit_economics_sku"],
        "where": "Заполняется вручную по шаблону «Классификатор SKU / Категория / Бренд»",
        "unavailable_if_missing": "Без него нельзя группировать аналитику по категориям и брендам.",
    },
    "manual_adjustments": {
        "title": "Ручные корректировки (ручная таблица)",
        "required": False,
        "group": GROUP_MANUAL,
        "supports_multiple": False,
        "blocks": ["finance"],
        "where": "Заполняется вручную по шаблону «Ручные корректировки»",
        "unavailable_if_missing": "Без него ручные корректировки сумм не будут применены к расчёту.",
    },
}

# Карта соответствия колонок: канонической считается ПЕРВАЯ подпись в списке.
# Остальные подписи — допустимые варианты названий из разных выгрузок Wildberries.
# Поиск колонок устойчив к пробелам, регистру и длинным официальным названиям
# (см. modules/columns.py: сначала точное совпадение, затем совпадение по вхождению).
COLUMN_MAPPING = {
    "finance_weekly": {
        "sku": ["Артикул поставщика", "Артикул", "Артикул WB", "Баркод"],
        "date": ["Дата продажи", "Дата заказа", "Дата операции", "Дата", "Период"],
        "amount": [
            "К перечислению Продавцу за реализованный Товар",
            "К перечислению за товар",
            "Сумма к перечислению",
            "К перечислению",
            "Итого к оплате",
        ],
        "logistics": [
            "Услуги по доставке товара покупателю",
            "Стоимость логистики",
            "Логистика",
            "Услуги по доставке",
        ],
        "storage": ["Стоимость хранения", "Хранение"],
        "penalty": ["Общая сумма штрафов", "Штрафы", "Удержания"],
    },
    "sales_report": {
        "sku": ["Артикул поставщика", "Артикул"],
        "date": ["Дата", "Дата продажи"],
        "qty": ["Количество", "Кол-во"],
        "amount": ["Сумма", "Выручка", "Сумма продаж"],
    },
    "stocks_report": {
        "sku": ["Артикул поставщика", "Артикул"],
        "warehouse": ["Склад", "Название склада"],
        "available": ["Доступно", "Остаток", "Доступно к заказу"],
        "in_transit": ["В пути", "Транзит"],
    },
}

# Человеко-понятные названия распознаваемых полей финансового отчёта (для диагностики).
FINANCE_FIELD_LABELS = {
    "sku": "Артикул",
    "date": "Дата",
    "amount": "Сумма к перечислению",
    "logistics": "Логистика",
    "storage": "Хранение",
    "penalty": "Штрафы и удержания",
}

# Минимально достаточный набор источников для каждого аналитического блока.
MIN_DATA_FOR_BLOCK = {
    "sales": ["finance_weekly"],
    "finance": ["finance_weekly"],
    "stocks": ["stocks_report"],
    "warehouses": ["stocks_report"],
    "unit_economics_sku": ["finance_weekly", "cost_price_manual"],
    "unit_economics_cabinet": ["finance_weekly"],
    "promotion": ["ads_report"],
    "trends": ["sales_report"],
    "risks": ["stocks_report", "turnover_report"],
}

# Человеко-понятные названия аналитических блоков (для итогового Excel и интерфейса).
BLOCK_LABELS = {
    "sales": "Продажи",
    "finance": "Финансы",
    "stocks": "Остатки на складах",
    "warehouses": "Склады",
    "unit_economics_sku": "Юнит-экономика по SKU",
    "unit_economics_cabinet": "Юнит-экономика кабинета",
    "promotion": "Продвижение и реклама",
    "trends": "Тренды продаж",
    "risks": "Риски и дефицит",
}
