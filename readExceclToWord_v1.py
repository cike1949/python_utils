#!/usr/bin/python
# -*- coding: UTF-8 -*-

# ******************************************************
# Author       : Ruopeng Huang
# Last modified: 2024-10-13 19:54
# Email        : cike1949@gmail.com
# Filename     : readExceclToWord.py
# Description  : v0.1
# ******************************************************

import sys
import os
import openpyxl
from datetime import datetime
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt,RGBColor

def replaceWordfromExcel(file_excel):
	# load excel file
	workbook = openpyxl.load_workbook(file_excel, read_only=True, data_only=True)
	# get sheet by name
	sheetEmailSummary = "Email Summary"
	sheetWeeklyReport = "周报"
	ws_ES = workbook[sheetEmailSummary]
	ws_WR = workbook[sheetWeeklyReport]
	# get cells values
	date_str = ws_ES['H2'].value
	gross_profit_margin = ws_WR['E8'].value
	growth_rate = ws_WR['G8'].value
	gross_profit_monthly = ws_WR['D8'].value
	# generate formated strings to replace
	new_date = datetime.strptime(date_str, '%Y-%m-%d').date()
	update_month = str(new_date.month)
	update_day = str(new_date.day)
	update_year = str(new_date.year)
	update_gross_profit_margin  = '{:,.0f}'.format(gross_profit_margin)
	update_growth_rate = '{:.0%}'.format(growth_rate)
	update_gross_profit_monthly = '{:.2f}'.format(gross_profit_monthly)
	# print out to check
	print("update info:")
	print(update_month,'月',update_day,'日')
	print('贸易毛利率为',update_gross_profit_margin)
	print('较去年同期增加',update_growth_rate)
	print(update_month,'月月度毛利增加',update_gross_profit_monthly)
	print('仅为',update_year,'年贸易毛利')

	# load word file
	doc = Document()
	doc.styles['Normal'].font.name = u'宋体'
	doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
#	doc.styles['Normal'].font.name = u'等线'
#	doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'等线')
	doc.styles['Normal'].font.size = Pt(10.5)
	doc.styles['Normal'].font.color.rgb = RGBColor(0, 0, 0)
	str1 = '按照单元风控日报口径，截至{}月{}日,'.format(update_month, update_day)
	str2 = '单元当年累计贸易毛利为{}万美元，'.format(update_gross_profit_margin)
	str3 = '较去年同期增加{}。'.format(update_growth_rate)
	str4 = '{}月月度毛利增加{}万美元，主要为。同比变化情况如下。'.format(update_month, update_gross_profit_monthly)
	str5 = '（统计口径：仅为{}年贸易毛利）'.format(update_year)
	myParagraph = str1 + str2 + str3 + str4 + str5
	doc.add_paragraph(myParagraph)
	file_new_word = '周报 {}.{}.{}.docx'.format(update_year, update_month, update_day)
	doc.save(file_new_word)

	print('Hi 旭瑜, replacement work has been completed, created a new word file <%s>'%file_new_word)

def main():
	argv_len = len(sys.argv)
	if argv_len != 2:
		print("Error: please check the numbers of paramters in python command! number %d"%argv_len)
		return
	file_excel = sys.argv[1]
	if (os.path.splitext(file_excel)[-1] != ".xlsx") and (os.path.splitext(file_excel)[-1] != ".xls"):
		print("Error: please input correct Excel file(*.xlsx *.xls)!")
		return
	replaceWordfromExcel(file_excel)

if __name__ == "__main__":
	main()
