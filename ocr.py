import os
import cv2
import numpy as np
import argparse
import pdf2image
from PIL import Image
from paddleocr import PaddleOCR
import xlsxwriter

dateInfo_N = 'Exercice N'
dateInfo_N1 = 'Exercice N-1'


pdocr = PaddleOCR(use_angle_cls=True, lang='en', show_log = False)

sub_res_temp = pdocr.ocr("./1.jpg", cls=False)
for line in sub_res_temp[0]:
    print(line)

def check_key(str_data):
    str_data = str_data.strip()
    str_data = str_data[:8]
    res_b = str_data.isnumeric()
    if res_b == True :
       if(int(str_data) > 100000) :
           return True, int(str_data)

    return False, 0

def compare_key(str_data, compare_str):
    str_data = str_data.strip()
    compare_str = compare_str.strip()
    # pos_find = str_data.find(compare_str)
    if(compare_str == str_data ) :
        return True
    else:
        return False

def sub_image_proc(img, h_s, h_e, w_s, w_e, sub_dateInfo_list):

    sub_path = "./image/temp.png"
    sub_img = img[int(h_s) : int(h_e), int(w_s) : int(w_e)]
    cv2.imwrite(sub_path, sub_img)

    sub_res = pdocr.ocr(sub_path, cls=False)
    print('row : ')
    sub_dateInfo_list.append(sub_res[0])
    for line in sub_res[0]:
        print(line)

def get_num_info(data_info, w_s, w_e, w_org_pos):

    num_data = ''

    for i in range(len(data_info)):
        w_s_r = data_info[i][0][0][0] + w_org_pos
        w_e_r = data_info[i][0][2][0] + w_org_pos
        if w_s_r > w_s - 20 and w_e_r < w_e + 20:
            num_data += str(data_info[i][1][0])

    num_data = num_data.replace('-', '')
    num_data = num_data.replace(' ', '')
    num_data = num_data.strip()
    if(num_data == ''):
        return 0.0
    else:
        return float(num_data)

if __name__ == '__main__':

    # check argument
    # parser = argparse.ArgumentParser()
    # parser.add_argument('-i', type=str, default=None, help='input pdf')
    # parser.add_argument('-o', type=str, default='./result',  help='result pdf')
    # args = parser.parse_args()

    workbook = xlsxwriter.Workbook('./extracted.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Account Number')
    worksheet.write('B1', 'Account Title')
    worksheet.write('C1', 'Date')
    worksheet.write('D1', 'Amount')
    row = 1
    col = 0

    for img_idx in range(5):
        dateInfo_list = []
        sub_dateInfo_list = []
        mainkey_list = []
        dateInfo_data_list = []
        # extract data
        extracted_data = []

        img_path = './image/' + str(img_idx+1) + '.png'
        result = pdocr.ocr(img_path, cls=False)

        for idx in range(len(result)):
            res = result[idx]
            for line in res:
                if(compare_key(line[1][0] , dateInfo_N) == True):
                    dateInfo_list.append(line[0])

                if(compare_key(line[1][0] , dateInfo_N1) == True):
                    dateInfo_list.append(line[0])

                b_num, real_num = check_key(line[1][0])
                if(b_num == True) :
                    mainkey_list.append(line[0])
                    # print(real_num)

        # print(dateInfo_list)

        opencvImage = cv2.imread(img_path)

        # extracte date N, date N-1
        # date N
        w_s_N = dateInfo_list[0][0][0] - 35
        w_e_N = dateInfo_list[0][2][0] + 35
        h_s_N = dateInfo_list[0][0][1] + 25
        h_e_N = dateInfo_list[0][2][1] + 25
        sub_image_proc(opencvImage, h_s_N, h_e_N, w_s_N, w_e_N, dateInfo_data_list)
        # date N - 1
        w_s_N1 = dateInfo_list[1][0][0] - 20
        w_e_N1 = dateInfo_list[1][2][0] + 20
        h_s_N1 = dateInfo_list[1][0][1] + 25
        h_e_N1 = dateInfo_list[1][2][1] + 25
        sub_image_proc(opencvImage, h_s_N1, h_e_N1, w_s_N1, w_e_N1, dateInfo_data_list)

        print(dateInfo_data_list[0][0][1][0])
        print(dateInfo_data_list[1][0][1][0])

        # create sub image
        for idx in range(len(mainkey_list)):
            w_s = mainkey_list[idx][0][0]
            w_e = dateInfo_list[1][2][0] + 50
            h_s = mainkey_list[idx][0][1] - 5
            h_e = mainkey_list[idx][2][1] + 5
            sub_image_proc(opencvImage, h_s, h_e, w_s, w_e, sub_dateInfo_list)


        for idx in range(len(sub_dateInfo_list)):
            line_data = []
            line_data.append(sub_dateInfo_list[idx][0][1][0])
            line_data.append(sub_dateInfo_list[idx][1][1][0])
            num_1 = get_num_info(sub_dateInfo_list[idx], w_s_N, w_e_N, mainkey_list[idx][0][0])
            num_2 = get_num_info(sub_dateInfo_list[idx], w_s_N1, w_e_N1, mainkey_list[idx][0][0])

            line_data.append(num_1)
            line_data.append(num_2)

            extracted_data.append(line_data)


        for idx in range(len(extracted_data)):

            str_data = dateInfo_data_list[0][0][1][0]
            str_data = str_data[:10]

            str_data_1 = dateInfo_data_list[1][0][1][0]
            str_data_1 = str_data_1[:10]

            worksheet.write(row, col, extracted_data[idx][0])
            worksheet.write(row, col + 1, extracted_data[idx][1])
            worksheet.write(row, col + 2, str_data)
            worksheet.write(row, col + 3, extracted_data[idx][2])

            worksheet.write(row+1, col, extracted_data[idx][0])
            worksheet.write(row+1, col + 1, extracted_data[idx][1])
            worksheet.write(row+1, col + 2, str_data_1)
            worksheet.write(row+1, col + 3, extracted_data[idx][3])
            row += 2

    workbook.close()