from pathlib import Path

from data_scripts_module.data_handler import DataManager


BASE_DIR = Path(__file__).resolve().parent

def main():
    data_manager = DataManager()

    data_manager.database = str(BASE_DIR / "data" / "db" / "CarSalesData.db")
    data_manager.filename_ = str(BASE_DIR / "data" / "Training.xlsx")
    data_manager.all_db_to_excel()


if __name__ == "__main__":
    main()