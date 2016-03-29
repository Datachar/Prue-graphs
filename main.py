import os
import xlrd
import operator
import numpy as np
import pandas as pd
from textwrap import wrap
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec

name_graphs = ['Facebook/Page', 'Facebook/Likes', 'Facebook/Post',
               'Twitter/Page', 'Twitter/Tweets', 'Twitter/Followers',
               'Youtube/Account', 'Youtube/Subscribers', 'Youtube/Videos', 'Youtube/Views',
               'LinkedIn/Account', 'LinkedIn/Connections', 'LinkedIn/Posts']

title = {
    name_graphs[0]: 'Facebook Page For Marketing - Top 48 Placeholders',
    name_graphs[1]: 'Number of Page Likes - Top 48 Positions with Facebook Promotional Page',
    name_graphs[2]: 'Average Posts Per Month by Instructor - Top 48 Positions with Facebook Promotional Page',
    name_graphs[3]: 'Twitter Account for Marketing - Top 48 Positions',
    name_graphs[4]: 'Total Twitter Tweets - Top 48 Positions with Twitter Account',
    name_graphs[5]: 'Number of Followers - Top 48 Positions with Twitter Account',
    name_graphs[6]: 'YouTube Account for Marketing - Top 48 Positions',
    name_graphs[7]: 'Number of YouTube Subscribers - Top 48  Positions with YouTube Channel',
    name_graphs[8]: 'Number of YouTube Videos - Top 48 Positions with YouTube Channel',
    name_graphs[9]: 'Number of YouTube Channel Views - Top 48 Positions with YouTube channel',
    name_graphs[10]: 'LinkedIn Account - Top 48 Positions',
    name_graphs[11]: 'Number of Connections - Top 48 Positions with LinkedIn Account',
    name_graphs[12]: 'Number of Posts - Top 48  Positions with LinkedIn Account',

}
bar_colors = ['#EE5363', '#57ccc6', '#f2b354', '#7cc576', '#c7cf48']
label = {
    name_graphs[0]: ['Have a Personal Facebook Page for Marketing',
                     'Have a Promotional Facebook Page for Marketing',
                     'Do Not Have a Facebook Page for Marketing'
                     ],
    name_graphs[1]: ['0', '1 - 100', '101 - 1,000', '1,001 - 10,000', '> 10,000'],
    name_graphs[2]: ['0', '1 - 10', '11 - 20', '21 - 30', '> 30'],
    name_graphs[3]: ['Do Not Have Twitter Account',
                     'Have Twitter Account'],
    name_graphs[4]: ['0', '1 - 1,000', '1,001 - 10,000', '10,001 - 100,000', '> 100,000'],
    name_graphs[5]: ['0', '1 - 1,000', '1,001 - 10,000', '10,001 - 100,000', '> 100,000'],
    name_graphs[6]: ['Do Not Have YouTube Account',
                     'Have YouTube Account'],
    name_graphs[7]: ['0', '1 - 100', '101 - 1,000', '1,001 - 10,000', '> 10,000'],
    name_graphs[8]: ['0', '1 - 100', '101 - 300', '301 - 500', '> 500'],
    name_graphs[9]: ['0', '1 - 1,000', '1,001 - 10,000', '10,001 - 100,000', '> 100,000'],
    name_graphs[10]: ['Do Not Have LinkedIn Account',
                      'Have LinkedIn Account'],
    name_graphs[11]: ['0', '1 - 100', '101 - 300', '301 - 500', '> 500'],
    name_graphs[12]: ['0', '1 - 10', '11 - 50', '51 - 100', '> 100'],
}
width_column = 1
y_step = np.arange(0, 100, 10)
subcategory_colors = ['gray', 'black']


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
        '/home/misha/Downloads/Data Sets 24-03/Web Development.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Office Productivity_Apple.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Office Productivity_Microsoft.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Office Productivity_Google.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Office Productivity_SAP.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Office Productivity_Oracle.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Office Productivity_Intuit.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Office Productivity_Other.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Office Productivity_Salesforce.xlsx',
        '/home/misha/Downloads/Data Sets 24-03/Development_Web Development.xlsx'
    ]
    header_data = create_header(file_paths)
    dfs = [pd.read_excel(io=file_path, skiprows=6) for file_path in file_paths]
    for i, key in enumerate(label.keys()):
        print(i + 1, '.', key)
        title[key] = '\n'.join(wrap(title[key], 50))
        y_label = ['\n'.join(wrap(l, 18)) for l in label[key]]
        if not os.path.exists(key):
            os.makedirs(key)
        draw_all_category_into_single_file(dfs, header_data, y_label, key)
        draw_all_category_into_separate_files(dfs, header_data, y_label, key)
        draw_average_by_categories_into_single_file(dfs, header_data, y_label, key)


def draw_all_category_into_single_file(df, header_data, y_label, key):
    plt.close('all')
    all_category = all_file_category(header_data)
    count_category = len(all_category)
    width_image = len(df)
    gs = gridspec.GridSpec(len(y_label) + 1, width_image)
    size_2 = 0
    for count in range(1, count_category + 1):
        data = create_data_for_category(df, header_data, all_category[count - 1], key)
        y = [sorted(val.items(), key=operator.itemgetter(0)) for val in data]
        plt.subplots_adjust(hspace=.001)
        x_step = np.arange(len(data))
        size_ = len(y)
        for i in range(len(y_label)):
            data = [el[i][1] for el in y]
            free_data = [(100 - el[i][1]) for el in y]
            plt.subplot(gs[i, size_2 - 3 * (count - 1): size_2 + size_ - 3 * (count - 1)])
            plt.subplots_adjust(hspace=.001, wspace=0.05)
            if count == 1:
                plt.ylabel(y_label[i], labelpad=106 + len(df), rotation='horizontal',
                           horizontalalignment='left', color=bar_colors[i], size=15)
            plt.bar(x_step, data, width_column, color=bar_colors[i], label='787')
            text = ["{:10.1f}%".format(d) for d in data]
            for j in range(len(data)):
                height_percent = data[j] - 5
                if data[j] < 8:
                    height_percent += 10
                plt.text(x_step[j] + width_column / 4, height_percent, text[j], horizontalalignment='center',
                         verticalalignment='center', color='black', weight='bold', size=11)
            plt.bar(x_step, free_data, width_column, color='w', bottom=data)
            plt.xticks(x_step, '')
            plt.yticks(y_step, '')
        name_subcategory = [data['Subcategory'] for data in header_data if
                            str(data['Category']) == str(all_category[count - 1])]
        plt.xticks(x_step + width_column / 2, name_subcategory, rotation=90,
                   size=14, color=subcategory_colors[count % len(subcategory_colors)])
        plt.xlabel(all_category[count - 1], weight='bold',
                   size=20, color=subcategory_colors[count % len(subcategory_colors)])
        size_2 += size_ + 3
    fig = plt.gcf()
    fig.set_size_inches(width_image, len(y_label) * 4)
    plt.savefig(key + '/' + 'All_graphs.png', dpi=150)
    print("IMAGE SAVE: %s/All graphs.png" % key)


def draw_all_category_into_separate_files(df, header_data, y_label, key):
    plt.close('all')
    all_category = all_file_category(header_data)
    count_category = len(all_category)
    for count in range(1, count_category + 1):
        data = create_data_for_category(df, header_data, all_category[count - 1], key)
        width_image = len(data)
        fig = plt.figure()
        fig.set_size_inches(width_image + 3, len(y_label) * 4)
        y = [sorted(val.items(), key=operator.itemgetter(0)) for val in data]
        fig.subplots_adjust(hspace=.001)
        x_step = np.arange(len(data))
        gs = gridspec.GridSpec(len(y_label) + 1, width_image + 4)
        for i in range(len(y_label)):
            data = [el[i][1] for el in y]
            free_data = [(100 - el[i][1]) for el in y]
            fig.add_subplot(gs[i, 2:-2])
            if i == 0:
                plt.title(title[key], weight='bold', size=15)
            plt.ylabel(y_label[i], labelpad=125, rotation='horizontal',
                       horizontalalignment='left', color=bar_colors[i])
            plt.bar(x_step, data, width_column, color=bar_colors[i], label='787')
            text = ["{:10.1f}%".format(d) for d in data]
            for j in range(len(data)):
                height_percent = data[j] - 5
                if data[j] < 8:
                    height_percent += 10
                plt.text(x_step[j] + width_column / 4, height_percent, text[j], horizontalalignment='center',
                         verticalalignment='center', color='black', weight='bold', size=11)
            plt.bar(x_step, free_data, width_column, color='w', bottom=data)
            plt.xticks(x_step, '')
            plt.yticks(y_step, '')
        name_subcategory = [data['Subcategory'] for data in header_data
                            if str(data['Category']) == str(all_category[count - 1])]
        plt.xticks(x_step + width_column / 2, name_subcategory, rotation=90, size=14)
        plt.xlabel(all_category[count - 1], weight='bold', size=20)
        print('Image save: %s/%s' % (key, all_category[count - 1]), '.png')
        fig.savefig(key + '/' + all_category[count - 1] + '.png', dpi=80)


def draw_average_by_categories_into_single_file(df, header_data, y_label, key):
    all_category = all_file_category(header_data)
    count_category = len(all_category)
    width_image = count_category
    gs = gridspec.GridSpec(len(y_label) + 1, width_image + 4)
    y = []
    for k in range(1, count_category + 1):
        data = create_data_for_category(df, header_data, all_category[k - 1], key)
        average_data = []
        for key_ in data[0].keys():
            average = sum(d[key_] for d in data) / len(data)
            average_data.append((key_, average))
        y.append(sorted(average_data))
    plt.subplots_adjust(hspace=.001)
    x_step = np.arange(len(y))
    for i in range(len(y_label)):
        data = [el[i][1] for el in y]
        free_data = [(100 - el[i][1]) for el in y]
        plt.subplot(gs[i, 2:-2])
        if i == 0:
            plt.title(title[key], weight='bold', size=len(y) + 10)
        plt.ylabel(y_label[i], labelpad=125, rotation='horizontal', horizontalalignment='left',
                   color=bar_colors[i])
        plt.bar(x_step, data, width_column, color=bar_colors[i], label='787')
        text = ["{:10.1f}%".format(d) for d in data]
        for j in range(len(data)):
            height_percent = data[j] - 5
            if data[j] < 8:
                height_percent += 10
            plt.text(x_step[j] + width_column / 4, height_percent, text[j], horizontalalignment='center',
                     verticalalignment='center', color='black', weight='bold', size=11)
        plt.bar(x_step, free_data, width_column, color='w', bottom=data)
        plt.xticks(x_step, '')
        plt.yticks(y_step, '')
    name_subcategory = [data for data in all_category]
    plt.xticks(x_step + width_column / 2, name_subcategory, rotation=90, size=14)
    plt.xlabel('Average all category', weight='bold', size=20)
    fig = plt.gcf()
    fig.set_size_inches(width_image + 5, len(y_label) * 4)
    plt.savefig(key + '/Average_all_graphs.png', dpi=150)
    print("IMAGE SAVE: %s/Average all graphs.png" % key)


def create_data_for_category(df, header_data, category, key):
    data = []
    for i, df_ in enumerate(df):
        if header_data[i]['Category'] == category:
            if key == name_graphs[0]:
                d = fb_page(df_)
            elif key == name_graphs[1]:
                d = fb_likes(df_)
            elif key == name_graphs[2]:
                d = fb_average_post(df_)
            elif key == name_graphs[3]:
                d = twitter_page(df_)
            elif key == name_graphs[4]:
                d = twitter_tweets(df_)
            elif key == name_graphs[5]:
                d = twitter_followers(df_)
            elif key == name_graphs[6]:
                d = youtube_account(df_)
            elif key == name_graphs[7]:
                d = youtube_subscribers(df_)
            elif key == name_graphs[8]:
                d = youtube_videos(df_)
            elif key == name_graphs[9]:
                d = youtube_views(df_)
            elif key == name_graphs[10]:
                d = youtube_views(df_)
            elif key == name_graphs[11]:
                d = youtube_views(df_)
            elif key == name_graphs[12]:
                d = youtube_views(df_)
            data.append(d)
    return data


def fb_page(df):
    data = df['Personal Facebook page?']
    size_ = df['Position on page 1'].count()
    not_have = (size_ - data.count())
    have_personal_page = len([i for i in data if i == 'Y'])
    have_promotion_page = size_ - not_have - have_personal_page
    not_have /= size_
    have_personal_page /= size_
    have_promotion_page /= size_
    data = {
        label[name_graphs[0]][0]: have_personal_page * 100,
        label[name_graphs[0]][1]: have_promotion_page * 100,
        label[name_graphs[0]][2]: not_have * 100
    }
    return data


def fb_likes(df):
    data = df['FB likes']
    size_ = df['Position on page 1'].count()
    percent = 100
    not_have = (size_ - data.count()) / size_
    have_1_100 = len([i for i in data if 0 < i <= 100]) / size_
    have_101_1000 = len([i for i in data if 101 <= i <= 1000]) / size_
    have_1001_10000 = len([i for i in data if 1001 <= i <= 10000]) / size_
    have_more_10000 = len([i for i in data if 10000 < i]) / size_
    data = {
        label[name_graphs[1]][0]: not_have * percent,
        label[name_graphs[1]][1]: have_1_100 * percent,
        label[name_graphs[1]][2]: have_101_1000 * percent,
        label[name_graphs[1]][3]: have_1001_10000 * percent,
        label[name_graphs[1]][4]: have_more_10000 * percent
    }
    return data


def fb_average_post(df):
    data = df['posts per month']
    size_ = df['Position on page 1'].count()
    percent = 100
    not_have = (size_ - data.count()) / size_
    have_1_10 = len([i for i in data if 0 < i <= 10]) / size_
    have_11_20 = len([i for i in data if 11 <= i <= 20]) / size_
    have_21_30 = len([i for i in data if 21 <= i <= 30]) / size_
    have_more_30 = len([i for i in data if 30 < i]) / size_
    data = {
        label[name_graphs[2]][0]: not_have * percent,
        label[name_graphs[2]][1]: have_1_10 * percent,
        label[name_graphs[2]][2]: have_11_20 * percent,
        label[name_graphs[2]][3]: have_21_30 * percent,
        label[name_graphs[2]][4]: have_more_30 * percent
    }
    return data


def twitter_page(df):
    data = df['Twitter']
    size_ = df['Position on page 1'].count()
    percent = 100
    have_personal_page = data.count()
    not_have = (size_ - have_personal_page) / size_
    have_personal_page /= size_
    data = {
        label[name_graphs[3]][0]: have_personal_page * percent,
        label[name_graphs[3]][1]: not_have * percent
    }
    return data


def twitter_tweets(df):
    data = df['Tweets']
    size_ = df['Position on page 1'].count()
    percent = 100
    not_have = (size_ - data.count()) / size_
    have_1_1000 = len([i for i in data if 0 < i <= 1000]) / size_
    have_1001_10000 = len([i for i in data if 1001 <= i <= 10000]) / size_
    have_10001_100000 = len([i for i in data if 10001 <= i <= 100000]) / size_
    have_more_100000 = len([i for i in data if 100000 < i]) / size_
    data = {
        label[name_graphs[4]][0]: not_have * percent,
        label[name_graphs[4]][1]: have_1_1000 * percent,
        label[name_graphs[4]][2]: have_1001_10000 * percent,
        label[name_graphs[4]][3]: have_10001_100000 * percent,
        label[name_graphs[4]][4]: have_more_100000 * percent
    }
    return data


def twitter_followers(df):
    data = df['Followers']
    size_ = df['Position on page 1'].count()
    percent = 100
    not_have = (size_ - data.count()) / size_
    have_1_1000 = len([i for i in data if 0 < i <= 1000]) / size_
    have_1001_10000 = len([i for i in data if 1001 <= i <= 10000]) / size_
    have_10001_100000 = len([i for i in data if 10001 <= i <= 100000]) / size_
    have_more_100000 = len([i for i in data if 100000 < i]) / size_
    data = {
        label[name_graphs[5]][0]: not_have * percent,
        label[name_graphs[5]][1]: have_1_1000 * percent,
        label[name_graphs[5]][2]: have_1001_10000 * percent,
        label[name_graphs[5]][3]: have_10001_100000 * percent,
        label[name_graphs[5]][4]: have_more_100000 * percent
    }
    return data


def youtube_account(df):
    data = df['Youtube']
    size_ = df['Position on page 1'].count()
    percent = 100
    have_personal_page = data.count()
    not_have = (size_ - have_personal_page) / size_
    have_personal_page /= size_
    data = {
        label[name_graphs[6]][0]: have_personal_page * percent,
        label[name_graphs[6]][1]: not_have * percent
    }
    return data


def youtube_subscribers(df):
    data = df['Youtube Subscribers']
    size_ = df['Position on page 1'].count()
    not_have = (size_ - data.count()) / size_
    data = [0 if str(i) == 'nan' else i for i in data]
    data = [float(i) if not isinstance(i, str) else int(i.replace(',', '')) for i in data]
    percent = 100
    have_1_100 = len([i for i in data if 0 < i <= 100]) / size_
    have_101_1000 = len([i for i in data if 101 <= i <= 1000]) / size_
    have_1001_10000 = len([i for i in data if 1001 <= i <= 10000]) / size_
    have_more_10000 = len([i for i in data if 10000 < i]) / size_
    data = {
        label[name_graphs[7]][0]: not_have * percent,
        label[name_graphs[7]][1]: have_1_100 * percent,
        label[name_graphs[7]][2]: have_101_1000 * percent,
        label[name_graphs[7]][3]: have_1001_10000 * percent,
        label[name_graphs[7]][4]: have_more_10000 * percent
    }
    return data


def youtube_videos(df):
    data = df['Youtube Videos']
    size_ = df['Position on page 1'].count()
    percent = 100
    not_have = (size_ - data.count()) / size_
    data = [0 if str(i) == 'nan' else i for i in data]
    data = [float(i) if not isinstance(i, str) else int(i.replace(',', '')) for i in data]
    have_1_100 = len([i for i in data if 0 < i <= 100]) / size_
    have_101_300 = len([i for i in data if 101 <= i <= 300]) / size_
    have_301_500 = len([i for i in data if 301 <= i <= 500]) / size_
    have_more_500 = len([i for i in data if 500 < i]) / size_
    data = {
        label[name_graphs[8]][0]: not_have * percent,
        label[name_graphs[8]][1]: have_1_100 * percent,
        label[name_graphs[8]][2]: have_101_300 * percent,
        label[name_graphs[8]][3]: have_301_500 * percent,
        label[name_graphs[8]][4]: have_more_500 * percent
    }
    return data


def youtube_views(df):
    data = df['Youtube Subscribers']
    size_ = df['Position on page 1'].count()
    percent = 100
    not_have = (size_ - data.count()) / size_
    data = [0 if str(i) == 'nan' else i for i in data]
    data = [float(i) if not isinstance(i, str) else int(i.replace(',', '')) for i in data]
    have_1_1000 = len([i for i in data if 0 < i <= 1000]) / size_
    have_1001_10000 = len([i for i in data if 1001 <= i <= 10000]) / size_
    have_10001_100000 = len([i for i in data if 10001 <= i <= 100000]) / size_
    have_more_100000 = len([i for i in data if 100000 < i]) / size_
    data = {
        label[name_graphs[9]][0]: not_have * percent,
        label[name_graphs[9]][1]: have_1_1000 * percent,
        label[name_graphs[9]][2]: have_1001_10000 * percent,
        label[name_graphs[9]][3]: have_10001_100000 * percent,
        label[name_graphs[9]][4]: have_more_100000 * percent
    }
    return data


def linked_in_account(df):
    data = df['Linkedin']
    size_ = df['Position on page 1'].count()
    percent = 100
    have_personal_page = data.count()
    not_have = (size_ - have_personal_page) / size_
    have_personal_page /= size_
    data = {
        label[name_graphs[10]][0]: have_personal_page * percent,
        label[name_graphs[10]][1]: not_have * percent
    }
    return data


def linked_in_connections(df):
    data = df['Connections']
    size_ = df['Position on page 1'].count()
    percent = 100
    not_have = (size_ - data.count()) / size_
    have_1_100 = len([i for i in data if 0 < i <= 100]) / size_
    have_101_300 = len([i for i in data if 101 <= i <= 300]) / size_
    have_301_500 = len([i for i in data if 301 <= i <= 500]) / size_
    have_more_500 = len([i for i in data if 500 < i]) / size_
    data = {
        label[name_graphs[11]][0]: not_have * percent,
        label[name_graphs[11]][1]: have_1_100 * percent,
        label[name_graphs[11]][2]: have_101_300 * percent,
        label[name_graphs[11]][3]: have_301_500 * percent,
        label[name_graphs[11]][4]: have_more_500 * percent
    }
    return data


def linked_in_posts(df):
    data = df['Posts']
    size_ = df['Position on page 1'].count()
    percent = 100
    not_have = (size_ - data.count()) / size_
    have_1_10 = len([i for i in data if 0 < i <= 10]) / size_
    have_11_30 = len([i for i in data if 11 <= i <= 30]) / size_
    have_31_50 = len([i for i in data if 31 <= i <= 50]) / size_
    have_more_50 = len([i for i in data if 50 < i]) / size_
    data = {
        label[name_graphs[12]][0]: not_have * percent,
        label[name_graphs[12]][1]: have_1_10 * percent,
        label[name_graphs[12]][2]: have_11_30 * percent,
        label[name_graphs[12]][3]: have_31_50 * percent,
        label[name_graphs[12]][4]: have_more_50 * percent
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


def all_file_category(header_data):
    all_data = []
    for i in header_data:
        if str(i['Category']) not in all_data:
            all_data.append(i['Category'])
    return all_data


if __name__ == '__main__':
    main()
