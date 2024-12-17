import pandas as pd
import os
from datetime import datetime
import configparser
import logging


def initialize_config():
    """讀取配置文件並檢查必要的資料夾。"""
    config = configparser.ConfigParser()
    config.read("config.ini", encoding="utf-8")

    # 獲取共用設定
    log_folder = config["General"]["log_folder"]
    output_folder = config["General"]["output_folder"]

    # 確保資料夾存在
    for folder in [log_folder, output_folder]:
        if not os.path.exists(folder):
            os.makedirs(folder)

    return config, log_folder, output_folder


def setup_logging(log_folder):
    """設定日誌系統。"""
    log_filename = os.path.join(log_folder, datetime.now().strftime("%Y-%m-%d.log"))
    logging.basicConfig(
        filename=log_filename,
        level=logging.ERROR,
        format="%(asctime)s - %(levelname)s - %(message)s",
        encoding="utf-8",
    )


def read_excel_data(file_path, sheet_name, columns):
    """讀取指定 Excel 文件中的特定欄位。"""
    try:
        columns_list = columns.split(",")  # 支援多欄位
        data = pd.read_excel(file_path, sheet_name=sheet_name, usecols=columns_list)
        return data
    except Exception as e:
        logging.error(f"讀取 Excel 文件失敗: {file_path}, 錯誤: {e}", exc_info=True)
        print(f"讀取 Excel 文件失敗: {file_path}, 錯誤: {e}")
        return pd.DataFrame()


def map_and_update(data_1, data_2, update_column):
    """基於 item_no 和 item_name (中文) 對比，更新並計算未匹配記錄來源。"""
    try:
        # 使用 item_no 和 item_name (中文) 進行合併匹配
        merged_data = pd.merge(
            data_2,
            data_1,
            on=["item_no", "item_name (中文)"],
            how="left",
            suffixes=("", "_from_files1"),
        )

        # 更新 Files_2 的對應列
        if update_column in merged_data.columns:
            merged_data[update_column] = merged_data[f"{update_column}_from_files1"]
            merged_data.drop(columns=[f"{update_column}_from_files1"], inplace=True)

        # 未匹配到 Files_1 的記錄
        unmatched_in_files1 = merged_data[merged_data[update_column].isna()]
        unmatched_in_files1 = unmatched_in_files1[["item_no", "item_name (中文)"]]

        # 未匹配到 Files_2 的記錄
        unmatched_in_files2 = data_1.merge(
            data_2, on=["item_no", "item_name (中文)"], how="left", indicator=True
        )
        unmatched_in_files2 = unmatched_in_files2[unmatched_in_files2["_merge"] == "left_only"]
        unmatched_in_files2 = unmatched_in_files2[["item_no", "item_name (中文)"]]

        return merged_data, unmatched_in_files1, unmatched_in_files2

    except Exception as e:
        logging.error(f"Mapping 過程發生錯誤: {e}", exc_info=True)
        print(f"Mapping 過程發生錯誤: {e}")
        return data_2, pd.DataFrame(), pd.DataFrame()


def main():
    """主程式邏輯。"""
    try:
        # 初始化配置
        config, log_folder, output_folder = initialize_config()

        # 設定日誌系統
        setup_logging(log_folder)

        # 讀取 Files_1 的配置
        file_1_path = config["Files_1"]["file_path"]
        sheet_1_name = config["Files_1"]["sheet_name"]
        mapping_1_columns = config["Files_1"]["mapping_column_name"]
        update_1_column = config["Files_1"]["update_column_name"]

        # 讀取 Files_2 的配置
        file_2_path = config["Files_2"]["file_path"]
        sheet_2_name = config["Files_2"]["sheet_name"]
        mapping_2_columns = config["Files_2"]["mapping_column_name"]
        update_2_column = config["Files_2"]["update_column_name"]

        # 讀取資料
        print("正在讀取 Files_1 的資料...")
        data_1 = read_excel_data(file_1_path, sheet_1_name, mapping_1_columns + "," + update_1_column)
        print("Files_1 的資料已讀取。")

        print("正在讀取 Files_2 的資料...")
        data_2 = read_excel_data(file_2_path, sheet_2_name, mapping_2_columns + "," + update_2_column)
        print("Files_2 的資料已讀取。")

        # 執行 Mapping 和更新操作
        print("正在進行 Mapping 和更新操作...")
        updated_data, unmatched_in_files1, unmatched_in_files2 = map_and_update(data_1, data_2, update_2_column)

        # 將結果保存到新文件
        output_file_path = os.path.join(output_folder, "updated_files2.xlsx")
        updated_data.to_excel(output_file_path, index=False)
        print(f"更新完成，結果已保存至: {output_file_path}")

        # 保存未匹配到 Files_1 的記錄
        if not unmatched_in_files1.empty:
            unmatched_in_files1_path = os.path.join(output_folder, "unmatched_in_files1.xlsx")
            unmatched_in_files1.to_excel(unmatched_in_files1_path, index=False)
            print(f"未匹配到 Files_1 的資料已保存至: {unmatched_in_files1_path}")

        # 保存未匹配到 Files_2 的記錄
        if not unmatched_in_files2.empty:
            unmatched_in_files2_path = os.path.join(output_folder, "unmatched_in_files2.xlsx")
            unmatched_in_files2.to_excel(unmatched_in_files2_path, index=False)
            print(f"未匹配到 Files_2 的資料已保存至: {unmatched_in_files2_path}")

    except Exception as e:
        logging.error(f"主程式執行時發生錯誤: {e}", exc_info=True)
        print(f"主程式執行時發生錯誤: {e}")


if __name__ == "__main__":
    # 執行主程式
    main()

    print("\nMapping 處理完成，請查看輸出結果！")
