"""
Централизованный реестр входных Excel-отчётов Wildberries и маппинг колонок.
Каждый отчёт описан статусом обязательности и связью с аналитическими блоками.
"""

REPORTS_REGISTRY = {
    "finance_weekly": {
        "title": "Еженедельный финансовый отчёт (детализация)",
        "required": True,
        "blocks": ["sales", "finance", "unit_economics_sku", "unit_economics_cabinet"],
        "where": "Главная -> Финансовые отчёты -> Еженедельные -> Скачать детализацию",
    },
    "sales_report": {
        "title": "Отчёт «Продажи»",
        "required": False,
        "blocks": ["sales", "trends"],
        "where": "Аналитика -> Отчёты -> Продажи",
    },
    "stocks_report": {
        "title": "Отчёт по остаткам на складах",
        "required": False,
        "blocks": ["stocks", "warehouses"],
        "where": "Аналитика -> Отчёты -> Остатки на складах",
    },
    "turnover_report": {
        "title": "Отчёт «Динамика оборачиваемости»",
        "required": False,
        "blocks": ["stocks", "risks"],
        "where": "Аналитика -> Отчёты -> Оборачиваемость",
    },
    "nomenclature_report": {
        "title": "Перечень номенклатур",
        "required": False,
        "blocks": ["sales", "unit_economics_sku"],
        "where": "Товары -> Карточки товаров -> Экспорт",
    },
    "storage_report": {
        "title": "Отчёт «Платное хранение»",
        "required": False,
        "blocks": ["finance", "unit_economics_sku"],
        "where": "Аналитика -> Отчёты -> Платное хранение",
    },
    "ads_report": {
        "title": "История затрат на продвижение",
        "required": False,
        "blocks": ["promotion", "unit_economics_sku"],
        "where": "Продвижение -> Реклама -> Финансы -> История затрат",
    },
    "cost_price_manual": {
        "title": "Ручная таблица себестоимости (вводится пользователем)",
        "required": False,
        "blocks": ["unit_economics_sku", "unit_economics_cabinet"],
        "where": "Заполняется вручную пользователем",
    },
}

COLUMN_MAPPING = {
    "finance_weekly": {
        "sku": ["Артикул", "Артикул поставщика", "Артикул WB"],
        "date": ["Дата продажи", "Дата", "Период"],
        "amount": ["К перечислению", "Сумма к перечислению", "Итого к оплате"],
        "logistics": ["Логистика", "Стоимость логистики"],
        "storage": ["Хранение", "Стоимость хранения"],
        "penalty": ["Штрафы", "Удержания"],
    },
    "sales_report": {
        "sku": ["Артикул", "Артикул поставщика"],
        "date": ["Дата", "Дата продажи"],
        "qty": ["Количество", "Кол-во"],
        "amount": ["Сумма", "Выручка"],
    },
    "stocks_report": {
        "sku": ["Артикул", "Артикул поставщика"],
        "warehouse": ["Склад", "Название склада"],
        "available": ["Доступно", "Остаток"],
        "in_transit": ["В пути", "Транзит"],
    },
}

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
