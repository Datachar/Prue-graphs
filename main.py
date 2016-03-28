import xlrd
import operator
import numpy as np
import pandas as pd
from textwrap import wrap
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec


def main():
    file_paths = [
        '/home/misha/Downloads/Data Sets 24-03/Business Law.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Communications.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Data & Analytics.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Databases.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Development Tools.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/E-Commerce.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Entrepreneurship.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Finance.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Game Development.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Hardware.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Home Business.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Human Resources.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Industry.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Intuit.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/IT Certification.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Management.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Media.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Mobile Apps.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Motivation.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Network & Security.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Operating Systems.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Operations.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Other.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Programming Languages.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Project Management.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Real Estate.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Salesforce.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Sales.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Self Esteem.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Software Engineering.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Software Testing.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Strategy.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Stress Management.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Web Development.xlsx'
    ]
    header_data = create_header(file_paths)
    dfs = [pd.read_excel(io=file_path, skiprows=6) for file_path in file_paths]
    all_data = [fb_page_for_marketing(df) for df in dfs]
    all_draw(dfs, header_data, all_data)
    draw_category(dfs, header_data)


def all_file_category(header_data):
    all_data = []
    for i in header_data:
        if str(i['Category']) not in all_data:
            all_data.append(i['Category'])
    return all_data


def create_data_for_category(df, header_data, category):
    data = []
    for i, df_ in enumerate(df):
        if header_data[i]['Category'] == category:
            d = fb_page_for_marketing(df_)
            data.append(d)
    return data


def fb_page_for_marketing(df):
    fb_page = df['Personal Facebook page?']
    size_ = df['Position on page 1'].count()
    not_have = (size_ - fb_page.count())
    have_personal_page = len([i for i in fb_page if i == 'Y'])
    have_promotion_page = size_ - not_have - have_personal_page
    not_have /= size_
    have_personal_page /= size_
    have_promotion_page /= size_
    data = {
        'Have a Personal Facebook Page For Marketing': have_personal_page * 100,
        'Have a Promotional Facebook Page For Marketing': have_promotion_page * 100,
        'Do Not Have a Facebook Page For Marketing': not_have * 100
    }
    return data


def create_header(file_path_):
    data_ = []
    for i in range(len(file_path_)):
        wb = xlrd.open_workbook(file_path_[i])
        wb = wb.sheet_by_index(0)
        data = {
            wb.cell_value(0, 0): wb.cell_value(0, 1),
            wb.cell_value(1, 0): wb.cell_value(1, 1),
            wb.cell_value(2, 0): wb.cell_value(2, 1),
            wb.cell_value(3, 0): wb.cell_value(3, 1),
            wb.cell_value(4, 0): wb.cell_value(4, 1)
        }
        data_.append(data)
    return data_


def all_draw(df, header_data, all_data):
    all_category = all_file_category(header_data)
    count = len(all_category)
    colors = ['#EE5363', '#F2B354', '#57CCC6']
    width_image = len(all_data)
    width = 1
    y_step = np.arange(0, 100, 10)
    gs = gridspec.GridSpec(4, width_image)
    size_2 = 0
    for k in range(1, count + 1):
        data = create_data_for_category(df, header_data, all_category[k - 1])
        y_label = sorted(data[0].items(), key=operator.itemgetter(0))
        y = [sorted(val.items(), key=operator.itemgetter(0)) for val in data]
        plt.subplots_adjust(hspace=.001)
        x_step = np.arange(len(data))
        size_ = len(y)
        y_label = ['\n'.join(wrap(l[0], 20)) for l in y_label]
        for i in range(3):
            data = [y[j][i][1] for j in range(len(y))]
            free_data = [(100 - y[j][i][1]) for j in range(len(y))]
            plt.subplot(gs[i, size_2-3*(k-1): size_2 + size_-3*(k-1)])
            plt.subplots_adjust(hspace=.001, wspace=0.02)
            if k == 1:
                plt.ylabel(y_label[i], labelpad=120, rotation='horizontal', horizontalalignment='left', color=colors[i])
            plt.bar(x_step, data, width, color=colors[i], label='787')
            text = ["{:10.1f}%".format(d) for d in data]
            for j in range(len(y)):
                plt.text(x_step[j] + width / 4, data[j] - 5, text[j], horizontalalignment='center',
                         verticalalignment='center', color='black', weight='bold', size=11)
            plt.bar(x_step, free_data, width, color='w', bottom=data)
            plt.xticks(x_step, )
            plt.yticks(y_step, '')
        name_subcategory = [data['Subcategory'] for data in header_data if
                            str(data['Category']) == str(all_category[k - 1])]
        plt.xticks(x_step + width / 2, name_subcategory, rotation=90, size=14)
        plt.xlabel(all_category[k - 1], weight='bold', size=20)
        size_2 += size_ + 3
    fig = plt.gcf()
    fig.set_size_inches(width_image, 10)
    plt.savefig('All graphs.png', dpi=150)
    print("IMAGE SAVE: All graphs.png")


def draw_category(df, header_data):
    all_category = all_file_category(header_data)
    count = len(all_category)
    colors = ['#EE5363', '#F2B354', '#57CCC6']
    width = 1
    y_step = np.arange(0, 100, 10)
    for k in range(1, count + 1):
        data = create_data_for_category(df, header_data, all_category[k - 1])
        width_image = len(data)
        fig = plt.figure()
        fig.set_size_inches(width_image+3, 10)
        y_label = sorted(data[0].items(), key=operator.itemgetter(0))
        y = [sorted(val.items(), key=operator.itemgetter(0)) for val in data]
        fig.subplots_adjust(hspace=.001)
        x_step = np.arange(len(data))
        gs = gridspec.GridSpec(4, width_image+4)
        y_label = ['\n'.join(wrap(l[0], 20)) for l in y_label]
        for i in range(3):
            data = [y[j][i][1] for j in range(len(y))]
            free_data = [(100 - y[j][i][1]) for j in range(len(y))]
            fig.add_subplot(gs[i, 2:-2])
            plt.ylabel(y_label[i], labelpad=120, rotation='horizontal', horizontalalignment='left', color=colors[i])
            plt.bar(x_step, data, width, color=colors[i], label='787')
            text = ["{:10.1f}%".format(d) for d in data]
            for j in range(len(y)):
                plt.text(x_step[j] + width / 4, data[j] - 5, text[j], horizontalalignment='center',
                         verticalalignment='center', color='black', weight='bold', size=11)
            plt.bar(x_step, free_data, width, color='w', bottom=data)
            plt.xticks(x_step, )
            plt.yticks(y_step, '')
        name_subcategory = [data['Subcategory'] for data in header_data if str(data['Category']) == str(all_category[k - 1])]
        plt.xticks(x_step + width / 2, name_subcategory, rotation=90, size=14)
        plt.xlabel(all_category[k - 1], weight='bold', size=20)
        print('Image save: ', all_category[k-1] + '.png')
        fig.savefig(all_category[k-1] + '.png', dpi=80)


if __name__ == '__main__':
    main()
