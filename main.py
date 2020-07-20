import MySQLdb
import json
import matplotlib.pyplot as plt
from os import path
import os
import csv
import xlsxwriter


def make_json_file():
    # User/Pass loaded from json config file so its safe for GitHub
    with open('config.txt') as cfg:
        login = json.load(cfg)

    conn = MySQLdb.connect(user=login['user'], passwd=login['pass'], db=login['db'], host=login['host'],
                           charset="utf8", use_unicode=True)
    cursor = conn.cursor()
    cursor.execute(""" select * from Contractor_USA """)
    rows = cursor.fetchall()
    cursor.close()
    print("Total rows found: " + str(len(rows)))
    with open('json.txt', 'w') as json_data:
        json.dump(rows, json_data)
    print('Data saved as json.txt')
    print('\n')


def read_json_file():
    with open('json.txt') as json_file:
        json_data = json.load(json_file)
    print("Total rows found: " + str(len(json_data)))
    return json_data


def data_analyze(data_in):

    data_category = {}
    data_cities = {}
    data_rating = {"Zero": 0, "One": 0, "Two": 0, "Three": 0, "Four": 0, "Five": 0}
    website_count = {"Website": 0, "None": 0}
    count_logger = 0
    for row in data_in:
        if count_logger % 10000 == 0:
            print("Progress: ", count_logger)
        count_logger += 1
        categories = row[7].split(",")
        for category in categories:
            if category not in data_category:
                data_category[category] = 1
            elif category in data_category:
                data_category[category] += 1
        city = row[3]
        if city not in data_cities:
            data_cities[city] = 1
        elif city in data_cities:
            data_cities[city] += 1

        # Local server row[11]
        # If online server row[12]
        rating = row[12]
        if rating:
            try:
                rating = int(rating)
            except ValueError:
                rating = float(rating)
            if 1.9 > rating > 1:
                data_rating["One"] += 1
            if 2.9 > rating > 2:
                data_rating["Two"] += 1
            if 3.9 > rating > 3:
                data_rating["Three"] += 1
            if 4.9 > rating > 4:
                data_rating["Four"] += 1
            if rating >= 5.0:
                data_rating["Five"] += 1
        else:
            data_rating["Zero"] += 1

        '''
        if rating not in data_rating:
            data_rating[rating] = 1
        else:
            data_rating[rating] += 1
        '''

        website = row[9]
        if website:
            website_count["Website"] += 1
        else:
            website_count["None"] += 1

    workbook = xlsxwriter.Workbook('data-analysis.xlsx')

    print("\n")

    print("Categories: ", len(data_category))
    sorted_categories = sort_data(data_category)
    # Makes csv file of top categories
    # make_csv(sorted_categories, 'categories')
    # Makes a pie chart of Categories
    make_plt_graph(sorted_categories,"Top 10 Categories", "categories", 10)
    # Add Sheet to Workbook
    make_xls(workbook, sorted_categories, 'categories.png', 'Categories')

    print("Cities: ", len(data_cities))
    sorted_cities = sort_data(data_cities)
    # Makes csv file of top cities
    # make_csv(sorted_cities, 'cities')
    # Makes a pie chart of Cities
    make_plt_graph(sorted_cities, "Top 10 Cities", "cities", 10)
    # Add Sheet to Workbook
    make_xls(workbook, sorted_cities, 'cities.png', 'Cities')

    print("Ratings: ", len(data_rating))
    sorted_ratings = sort_data(data_rating)
    # Makes csv file of ratings
    # make_csv(sorted_ratings, 'ratings')
    # Makes a pie chart of ratings
    make_plt_graph(sorted_ratings, "Top 10 Ratings", "ratings", 10)
    # Add Sheet to Workbook
    make_xls(workbook, sorted_ratings, 'ratings.png', 'Ratings')

    print("Number of Websites: ", website_count)
    sorted_website = sort_data(website_count)
    # Makes a pie chart of ratings
    make_plt_graph(sorted_website, "Listings with Websites", "websites", "")
    # Add Sheet to Workbook
    make_xls(workbook, sorted_website, 'websites.png', 'Websites')

    print("\n")

    # Option to write all data to workbook as well
    # make_xls(workbook, data_in, '', 'CSV Data')

    workbook.close()

    # Delete images after we close the workbook to avoid errors.
    os.remove("categories.png")
    os.remove("cities.png")
    os.remove("ratings.png")
    os.remove("websites.png")


def make_csv(list_in, file_name):
    with open(file_name + '.csv', 'w') as file_csv:
        w = csv.writer(file_csv, quoting=csv.QUOTE_ALL)
        for row in list_in:
            i = 0
            item = []
            while i < len(row):
                item.append(row[i])
                i += 1
            w.writerow(item)


def make_xls(workbook, list_in, image_in, sheet_name):

    worksheet = workbook.add_worksheet(sheet_name)
    i = 0
    print(len(list_in))

    for row in list_in:
        col = 0
        worksheet.set_column(col, len(list_in), 20)
        while col < len(row):
            item = row[col]
            worksheet.write_string(i, col, str(item))
            col += 1
        i += 1
    image_col = len(row)
    image_row = 1
    worksheet.insert_image(image_row, image_col, image_in)


def sort_data(dataset):
    # System for Top 10
    sorted_data = sorted(dataset.items(), key=lambda x: x[1], reverse=True)
    # top10Data = sortedCategory[0:10]
    # Print the top cities and categories out to console
    print(sorted_data)
    return sorted_data


def make_plt_graph(dataset, title, fname, top):

    # Create the a pie chart of data.
    plt_labels = []
    plt_values = []
    if top:
        for item in dataset[:top]:
            plt_labels.append(item[0])
            plt_values.append(item[1])
    else:
        for item in dataset:
            plt_labels.append(item[0])
            plt_values.append(item[1])
    plt.pie(plt_values, labels=plt_labels, autopct='%1.1f%%')
    plt.axis('equal')
    plt.title(title)
    plt.savefig(fname+'.png')
    plt.close()


if path.exists('json.txt'):
    print("Data file already exists. Loading...")
    data = read_json_file()
    print('Loaded Data from File')
    data_analyze(data)
    # Option to also export a csv
    # make_csv(data, 'data')
else:
    print('Data Missing, Creating JSON')
    make_json_file()
    data = read_json_file()
    data_analyze(data)
    # Option to also export a csv
    # make_csv(data, 'data')

print("Job Completed")
