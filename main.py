import xlrd
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from textwrap import wrap
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
    new_draw(dfs, header_data, all_data)


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
        'have a personal Facebook page for marketing': have_personal_page * 100,
        'have a promotional Facebook page for marketing': have_promotion_page * 100,
        'do not have a Facebook page for marketing': not_have * 100
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


def new_draw(df,header_datas, all_data):
    all_category = all_file_category(header_datas)
    count = len(all_category)
    colors = ['#EE5363', '#57CCC6', '#F2B354']
    width = 1
    y_step = np.arange(0, 100, 10)
    gs = gridspec.GridSpec(8, len(all_data)+2*len(all_category)+5)
    print(len(all_data)+3*len(all_category))
    size_2 = 0
    for k in range(1, count+1):
        datas = create_data_for_category(df, header_datas, all_category[k-1])
        ylabel = list(datas[0].keys())
        y = np.array([list(val.values()) for val in datas])
        plt.subplots_adjust(hspace=.001)
        x_step = np.arange(len(datas))
        ylabel = ['\n'.join(wrap(l, 20)) for l in ylabel]
        for i in range(3):
            data = [y[j][i] for j in range(len(y))]
            free_data = [(100 - y[j][i]) for j in range(len(y))]
            size_ = len(data)
            print(i, size_2, size_2+size_)
            plt.subplots_adjust(hspace=.001, wspace=0.45)
            plt.subplot(gs[i, size_2: size_2+size_])
            plt.ylabel(ylabel[i], rotation='horizontal', horizontalalignment='right')
            plt.bar(x_step, data, width, color=colors[i], label='787')
            text = ["{:10.1f}%".format(d) for d in data]
            for j in range(len(y)):
                plt.text(x_step[j] + width/4, data[j] - 5, text[j], horizontalalignment='center',
                         verticalalignment='center', color='black', weight='bold', size=11)
            plt.bar(x_step, free_data, width, color='w', bottom=data)
            plt.xticks(x_step,)
            plt.yticks(y_step, '')
        name_subcategory = [data['Subcategory'] for data in header_datas if str(data['Category']) == str(all_category[k-1])]
        plt.xticks(x_step + width / 2, name_subcategory, rotation=90, size=7)
        plt.xlabel(all_category[k-1], weight='bold', size=16)
        size_2 += size_ + 3
    ylabel = list(all_data[0].keys())
    y = np.array([list(val.values()) for val in all_data])
    plt.subplot(gs[5,:])
    plt.subplots_adjust(hspace=.001)
    plt.title('Facebook Page for Marketing', size=25, weight='heavy')
    x_step = np.arange(len(all_data))
    ylabel = ['\n'.join(wrap(l, 20)) for l in ylabel]
    for i in range(3):
        data = [y[j][i] for j in range(len(y))]
        free_data = [(100 - y[j][i]) for j in range(len(y))]
        plt.subplot(gs[i+5,4:40])
        plt.subplots_adjust(hspace=.001)
        plt.ylabel(ylabel[i], rotation='horizontal', horizontalalignment='right')
        plt.bar(x_step, data, width, color=colors[i], label='787')
        text = ["{:10.1f}%".format(d) for d in data]
        for i in range(len(y)):
            plt.text(x_step[i] + width / 2, data[i] - 5, text[i], horizontalalignment='center',
                     verticalalignment='center', color='black', weight='bold', size=12)
        plt.bar(x_step, free_data, width, color='w', bottom=data)
        plt.xticks(x_step, '')
        plt.yticks(y_step, '')
    name_subcategory = [data['Subcategory'] for data in header_datas]
    plt.xticks(x_step + width / 4, name_subcategory, rotation=90, size=7)
    plt.xlabel('All category',weight='bold', size=16)
    fig = plt.gcf()
    fig.set_size_inches(40, 25)
    plt.savefig('Graphs.png', dpi=150)
    #plt.show()


if __name__ == '__main__':
    main()
