import pandas as pd

def stage_2(df: pd.DataFrame):
    results = []
    addresses = pd.read_excel("./assets/Список_адресов.xlsx", engine="calamine",)["Уникальные адреса"].tolist()
    # print(df)
    for address in addresses:
        subset = df[df["Грузополучатель и его адрес:"] == address]

        if not subset.empty:
        # Собираем все номера через запятую
            invoice_list = ", ".join(subset["Счет-фактура №"].astype(str).unique())
        
            dates_list = ", ".join(subset["Дата"].astype(str).unique())
            # Считаем сумму именно для этого адреса
            total_for_addr = pd.to_numeric(subset["Всего к оплате (9)"], errors='coerce').sum()
            
            results.append({
                "Адрес": address,
                "Номера УПД": invoice_list,
                "Даты": dates_list,
                "Сумма": total_for_addr
            })
    # 2. Превращаем результат в финальную таблицу
    final_lookup = pd.DataFrame(results)
    final_lookup.to_excel("./assets/данные-по-адресам.xlsx")
  
