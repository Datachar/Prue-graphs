from __future__ import division

import os
import sys
import xlrd
import time
import operator
import numpy as np
import pandas as pd
from glob import glob
from logging import getLogger
from os.path import join
from textwrap import wrap
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from datachar_armory.os_utils.logging_utils import configure_file_and_stream_logger

name_graphs = [join('Facebook', 'Page'), join('Facebook', 'Likes'), join('Facebook', 'AVPost'),
               join('Twitter', 'Page'), join('Twitter', 'Tweets'), join('Twitter', 'Followers'),
               join('Youtube', 'Account'), join('Youtube', 'Subscribers'), join('Youtube', 'Videos'),
               join('Youtube', 'Views'),
               join('LinkedIn', 'Account'), join('LinkedIn', 'Connections'), join('LinkedIn', 'Posts')
               ]
configure_file_and_stream_logger()
logger = getLogger()
path_to_output_directory = join(sys.argv[2], time.strftime("%d-%m-%Y"))
count_ = 0
while os.path.exists(path_to_output_directory):
    if count_ == 0:
        count_ += 1
        path_to_output_directory = '%s(%s)' % (path_to_output_directory, count_)
    else:
        path_to_output_directory = path_to_output_directory[:-(len(str(count_ - 1)) + 2)]
        path_to_output_directory = '%s(%s)' % (path_to_output_directory, count_)
    count_ += 1
name_graphs = [join(path_to_output_directory, name) for name in name_graphs]
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
    logger.info('Read files from ' + sys.argv[1])
    file_paths = glob(join(sys.argv[1], '*.xlsx'))
    file_paths = [file_path for file_path in file_paths if not os.path.basename(file_path).startswith('~$')]
    if len(file_paths) == 0:
        logger.info('No single .xlsx file format in ' + sys.argv[1])
        return 0
    dfs = [pd.read_excel(io=file_path, skiprows=6) for file_path in file_paths]
    header_data, file_paths = create_header(file_paths)
    delete_empty_df(dfs, header_data)
    for i, key in enumerate(label.keys()):
        logger.info(str(i + 1) + '. ' + str(key))
        y_label = ['\n'.join(wrap(l, 16)) for l in label[key]]
        if not os.path.exists(key):
            os.makedirs(key)
        draw_all_category_into_single_file(dfs, header_data, y_label, key)
        title[key] = '\n'.join(wrap(title[key], 50))
        draw_all_category_into_separate_files(dfs, header_data, y_label, key)
        draw_average_by_categories_into_single_file(dfs, header_data, y_label, key)


def draw_all_category_into_single_file(df, header_data, y_label, key):
    all_category = all_file_category(header_data)
    count_category = len(all_category)
    width_image = len(df)
    gs = gridspec.GridSpec(len(y_label) + 2, width_image)
    size_2 = 0
    max_name_subcategory = max([len(data['Subcategory']) for data in header_data])
    plt.subplot(gs[0, :])
    plt.text(0.5 - len(title[key]) / 300, 0.14, title[key], weight='bold',
             verticalalignment='center', size=width_image / 1.5)
    plt.gca().axison = False
    for count in range(1, count_category + 1):
        data = create_data_for_category(df, header_data, all_category[count - 1], key)
        y = [sorted(val.items(), key=operator.itemgetter(0)) for val in data]
        plt.subplots_adjust(hspace=.001)
        x_step = np.arange(len(data))
        size_ = len(y)
        for i in range(len(y_label)):
            data = [el[i][1] for el in y]
            free_data = [(100 - el[i][1]) for el in y]
            plt.subplot(gs[i + 1, size_2 - 3 * (count - 1): size_2 + size_ - 3 * (count - 1)])
            plt.xticks(x_step, '')
            plt.yticks(y_step, '')
            plt.subplots_adjust(hspace=.001, wspace=0.15)
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
        name_subcategory = [data['Subcategory'] for data in header_data if
                            str(data['Category']) == str(all_category[count - 1])]
        name_subcategory = [(' ' * (max_name_subcategory - len(subcategory) + (count % 2) * 3) + subcategory) for
                            subcategory in name_subcategory]
        plt.xticks(x_step + width_column / 2, name_subcategory, rotation=90, fontproperties='monospace',
                   size=14, color=subcategory_colors[count % len(subcategory_colors)])
        plt.xlabel(all_category[count - 1], weight='bold',
                   size=20, color=subcategory_colors[count % len(subcategory_colors)])
        size_2 += size_ + 3
    fig = plt.gcf()
    fig.set_size_inches(width_image, len(y_label) * 5 + (max_name_subcategory - len(y_label)) / 8)
    plt.savefig(join(key, 'All_graphs.png'), dpi=150)
    logger.info("IMAGE SAVE: %s All graphs.png" % key)


def draw_all_category_into_separate_files(df, header_data, y_label, key):
    plt.close('all')
    max_name_subcategory = max([len(data['Subcategory']) for data in header_data])
    all_category = all_file_category(header_data)
    count_category = len(all_category)
    for count in range(1, count_category + 1):
        data = create_data_for_category(df, header_data, all_category[count - 1], key)
        width_image = len(data)
        fig = plt.figure()
        fig.set_size_inches(width_image + 3, len(y_label) * 4 + max_name_subcategory / 10)
        y = [sorted(val.items(), key=operator.itemgetter(0)) for val in data]
        fig.subplots_adjust(hspace=.001)
        x_step = np.arange(len(data))
        gs = gridspec.GridSpec(len(y_label) + 1, width_image + 4)
        for i in range(len(y_label)):
            data = [el[i][1] for el in y]
            free_data = [(100 - el[i][1]) for el in y]
            fig.add_subplot(gs[i, 2:-2])
            plt.xticks(x_step, '')
            plt.yticks(y_step, '')
            if i == 0:
                plt.title(title[key], weight='bold', size=len(data) + 7)
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
        name_subcategory = [data['Subcategory'] for data in header_data
                            if str(data['Category']) == str(all_category[count - 1])]
        plt.xticks(x_step + width_column / 2, name_subcategory, rotation=90, size=14)
        plt.xlabel(all_category[count - 1], weight='bold', size=20)
        fig.savefig(join(key, all_category[count - 1] + '.png'), dpi=90)
        logger.info('Image save: ' + join(key, all_category[count - 1] + '.png'))


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
        plt.xticks(x_step, '')
        plt.yticks(y_step, '')
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
    name_subcategory = [data for data in all_category]
    plt.xticks(x_step + width_column / 2, name_subcategory, rotation=90, size=14)
    plt.xlabel('Average all category', weight='bold', size=20)
    fig = plt.gcf()
    fig.set_size_inches(width_image + 5, len(y_label) * 4)
    plt.savefig(join(key, 'Average_all_graphs.png'), dpi=150)
    logger.info("IMAGE SAVE: %s Average all graphs.png" % key)


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


def exclusive_instructors(df, data):
    instructors = list(df['Position on page 1'])
    data = list(data)
    count = 0
    data_count = 0
    while count != len(instructors) - 1:
        if instructors[count] == instructors[count + 1]:
            if str(data[count]) != 'nan':
                data_count += 1
            if data[count] == data[count + 1]:
                instructors.pop(count + 1)
                data.pop(count + 1)
            count += 1
        else:
            count += 1
            if str(data[count]) != 'nan':
                data_count += 1
    count += 1
    return data, count, data_count


def fb_page(df):
    data, size_, data_count = exclusive_instructors(df, df['Personal Facebook page?'])
    have_personal_page = len([i for i in data if i == 'Y'])
    not_have = size_ - data_count
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
    data, size_, data_count = exclusive_instructors(df, df['FB likes'])
    percent = 100
    not_have = (size_ - data_count) / size_
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
    data, size_, data_count = exclusive_instructors(df, df['posts per month'])
    percent = 100
    not_have = (size_ - data_count) / size_
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
    data, size_, data_count = exclusive_instructors(df, df['Twitter'])
    percent = 100
    have_personal_page = data_count
    not_have = (size_ - have_personal_page) / size_
    have_personal_page /= size_
    data = {
        label[name_graphs[3]][0]: have_personal_page * percent,
        label[name_graphs[3]][1]: not_have * percent
    }
    return data


def twitter_tweets(df):
    data, size_, data_count = exclusive_instructors(df, df['Tweets'])
    percent = 100
    not_have = (size_ - data_count) / size_
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
    data, size_, data_count = exclusive_instructors(df, df['Followers'])
    percent = 100
    not_have = (size_ - data_count) / size_
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
    data, size_, data_count = exclusive_instructors(df, df['Youtube'])
    percent = 100
    have_personal_page = data_count
    not_have = (size_ - have_personal_page) / size_
    have_personal_page /= size_
    data = {
        label[name_graphs[6]][0]: have_personal_page * percent,
        label[name_graphs[6]][1]: not_have * percent
    }
    return data


def youtube_subscribers(df):
    data, size_, data_count = exclusive_instructors(df, df['Youtube Subscribers'])
    not_have = (size_ - data_count) / size_
    data = [0 if str(i) == 'nan' else i for i in data]
    data = [float(str(i).replace(',', '')) for i in data]
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
    data, size_, data_count = exclusive_instructors(df, df['Youtube Videos'])
    percent = 100
    not_have = (size_ - data_count) / size_
    data = [0 if str(i) == 'nan' else i for i in data]
    data = [float(str(i).replace(',', '')) for i in data]
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
    data, size_, data_count = exclusive_instructors(df, df['Youtube Subscribers'])
    percent = 100
    not_have = (size_ - data_count) / size_
    data = [0 if str(i) == 'nan' else i for i in data]
    data = [float(str(i).replace(',', '')) for i in data]
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
    data, size_, data_count = exclusive_instructors(df, df['Linkedin'])
    percent = 100
    have_personal_page = data_count
    not_have = (size_ - have_personal_page) / size_
    have_personal_page /= size_
    data = {
        label[name_graphs[10]][0]: have_personal_page * percent,
        label[name_graphs[10]][1]: not_have * percent
    }
    return data


def linked_in_connections(df):
    data, size_, data_count = exclusive_instructors(df, df['Connections'])
    percent = 100
    not_have = (size_ - data_count) / size_
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
    data, size_, data_count = exclusive_instructors(df, df['Posts'])
    size_ = df['Position on page 1'].count()
    percent = 100
    not_have = (size_ - data_count) / size_
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
    data = []
    new_file = []
    for file in file_path_:
        try:
            wb = xlrd.open_workbook(file)
            wb = wb.sheet_by_index(0)
            data_ = {
                wb.cell_value(0, 0): wb.cell_value(0, 1),
                wb.cell_value(1, 0): wb.cell_value(1, 1),
                wb.cell_value(2, 0): wb.cell_value(2, 1),
                wb.cell_value(3, 0): wb.cell_value(3, 1),
                wb.cell_value(4, 0): wb.cell_value(4, 1)
            }
            data.append(data_)
            new_file.append(file)
        except:
            logger.info('cannot read ' + file)
    return data, new_file


def all_file_category(header_data):
    all_data = []
    for i in header_data:
        if str(i['Category']) not in all_data:
            all_data.append(i['Category'])
    return all_data


def delete_empty_df(dfs, header_data):
    for i, df in enumerate(dfs):
        if len(df) == 0:
            logger.info(header_data[i]['Category'] + '_' + header_data[i]['Subcategory'] + ' not have data')
            dfs.pop(i)
            header_data.pop(i)


if __name__ == '__main__':
    main()
