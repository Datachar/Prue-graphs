import xlrd
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


def main():
    file_path = '/home/misha/Downloads/Data Sets 24-03/Business Law.xlsx'
    header_data = create_header(file_path)
    df = pd.read_excel(io=file_path, skiprows=6)
    fb_data = fb_page_for_marketing(df)
    print(fb_data)
    draw_pie(fb_data, 'Facebook Page for Marketing', header_data)


def fb_page_for_marketing(df):
    fb_page = df['Personal Facebook page?']
    size_ = df['Position on page 1'].count()
    not_have = (size_ - fb_page.count())
    have_personal_page = len([i for i in fb_page if i == 'Y'])
    have_promotion_page = size_ - not_have - have_personal_page
    print(size_, not_have, have_personal_page, have_promotion_page)
    not_have /= size_
    have_personal_page /= size_
    have_promotion_page /= size_
    data = {
        'have a personal Facebook page for marketing': have_personal_page,
        'have a promotional Facebook page for marketing': have_promotion_page,
        'do not have a Facebook page for marketing': not_have
    }
    return data


def create_header(file_path_):
    wb = xlrd.open_workbook(file_path_)
    wb = wb.sheet_by_index(0)
    data_ = {
        wb.cell_value(0, 0): wb.cell_value(0, 1),
        wb.cell_value(1, 0): wb.cell_value(1, 1),
        wb.cell_value(2, 0): wb.cell_value(2, 1),
        wb.cell_value(3, 0): wb.cell_value(3, 1),
        wb.cell_value(4, 0): wb.cell_value(4, 1)
    }

    return data_


def draw_pie(data, title, header_data):
    labels = list(data.keys())
    sizes = list(data.values())
    colors = ['yellowgreen', 'gold', 'lightskyblue', 'lightcoral']
    plt.pie(sizes, labels=labels, colors=colors,
            autopct='%1.1f%%', startangle=90)
    # Set aspect ratio to be equal so that pie is drawn as a circle.
    fig = plt.gcf()
    fig.set_size_inches(18.5, 10.5)
    plt.title(title)
    plt.axis('equal')
    plt.savefig(header_data['Subcategory'] + '.png', dpi=100)

if __name__ == '__main__':
    main()
