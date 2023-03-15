from config import input_file, output_file, image_file
from data_processing import read_input_file, drop_columns, sort_and_filter, format_columns
from excel_export import save_dataframe_to_excel
from plotting import create_pie_chart

def main():
    df = read_input_file(input_file)
    df = drop_columns(df, columns_to_remove)
    df = sort_and_filter(df)
    df = format_columns(df)

    writer = save_dataframe_to_excel(df, output_file)
    writer.close()

    create_pie_chart(df, image_file)

if __name__ == '__main__':
    main()
