import pandas as pd
import numpy as np

def process_waybills(posylka_path, pp_path, filial_path):
    try:
        # Читаем Посылку (с заголовками)
        df_posylka = pd.read_excel(posylka_path)
        
        # Читаем сводки без заголовков, чтобы искать по фиксированным номерам колонок
        df_pp = pd.read_excel(pp_path, sheet_name=0, header=None)
        df_filial = pd.read_excel(filial_path, sheet_name=0, header=None)
        
        # Очистка PS в Посылке (удаляем пробелы и '.0', если Excel прочитал как число)
        df_posylka['PS_str'] = df_posylka['PS'].astype(str).str.strip().str.replace('.0', '', regex=False)
        
        # Подготовка словаря Головы (ПП)
        # 2 - PS, 4 - Количество, 14 - Доставка
        df_pp['PS_str'] = df_pp[2].astype(str).str.strip().str.replace('.0', '', regex=False)
        pp_dict = df_pp.dropna(subset=['PS_str']).set_index('PS_str').to_dict('index')
        
        # Подготовка словаря Филиала
        # 4 - PS, 7 - Количество, 1 - Доставка
        df_filial['PS_str'] = df_filial[4].astype(str).str.strip().str.replace('.0', '', regex=False)
        filial_dict = df_filial.dropna(subset=['PS_str']).set_index('PS_str').to_dict('index')
        
        count_head = 0
        count_filial = 0

        # Сведение данных
        for idx, row in df_posylka.iterrows():
            ps = str(row['PS_str'])
            if ps == 'nan' or not ps:
                continue
                
            needs_update = pd.isna(row['По счету']) or pd.isna(row['Доставка'])
            
            if needs_update:
                if ps in pp_dict:
                    if pd.isna(row['По счету']):
                        df_posylka.at[idx, 'По счету'] = pp_dict[ps].get(4, np.nan)
                    if pd.isna(row['Доставка']):
                        df_posylka.at[idx, 'Доставка'] = pp_dict[ps].get(14, np.nan)
                    count_head += 1
                        
                elif ps in filial_dict:
                    if pd.isna(row['По счету']):
                        df_posylka.at[idx, 'По счету'] = filial_dict[ps].get(7, np.nan)
                    if pd.isna(row['Доставка']):
                        df_posylka.at[idx, 'Доставка'] = filial_dict[ps].get(1, np.nan)
                    count_filial += 1

        # Убираем вспомогательную колонку
        df_posylka.drop(columns=['PS_str'], inplace=True)
        
        # Текст для предпросмотра на экране (разделитель точка с запятой)
        preview_csv = df_posylka.to_csv(sep=';', index=False, na_rep='')
        
        return True, df_posylka, preview_csv, count_head, count_filial
    
    except Exception as e:
        return False, None, str(e), 0, 0
