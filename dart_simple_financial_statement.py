#-*- coding:utf-8 -*-
# Read text files
import os
import xlsxwriter
import urllib.request
from bs4 import BeautifulSoup
import zipfile
import sys
import re
from PyQt5.QtWidgets import *
from PyQt5 import uic
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient import discovery

def read_raw_data(file_name):
	cur_dir = os.getcwd()
	f_read = open(os.path.join(cur_dir, file_name), "r")

	raw_data = []

	line_number = 0
	#Read lines
	while True:
		line_number = line_number + 1
		line = f_read.readline()
		if not line: break

		#재무제표종류	
		#종목코드	
		#회사명	
		#시장구분	
		#업종	
		#업종명	
		#결산월	
		#결산기준일	
		#보고서종류	
		#통화	
		#항목코드	
		#항목명		
		#당기		
		#전기		
		#전전기		

		word_line = line.split('\t')
		raw_data.append(word_line)

	return raw_data

def scrape_cashflow_statement(raw_data, index, corp):

	cashflow_sub_list = {}

	cashflow_sub_list['CashFlowsFromUsedInOperatingActivities']		=	-1.0
	cashflow_sub_list['ProfitLossForStatementOfCashFlows']			=	-1.0
	cashflow_sub_list['AdjustmentsForReconcileProfitLoss']			=	-1.0
	cashflow_sub_list['AdjustmentsForDepreciationExpense']			=	-1.0
	cashflow_sub_list['CashFlowsFromUsedInInvestingActivities']		=	-1.0
	cashflow_sub_list['PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities']		=	-1.0
	cashflow_sub_list['PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities']				=	-1.0
	cashflow_sub_list['PurchaseOfInvestmentProperty']											=	-1.0
	cashflow_sub_list['ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities']	=	-1.0
	cashflow_sub_list['ProceedsFromSalesOfIntangibleAssetsClassifiedAsInvestingActivities']		=	-1.0
	cashflow_sub_list['ProceedsFromSalesOfInvestmentProperty']	=	-1.0
	cashflow_sub_list['CashFlowsFromUsedInFinancingActivities']	=	-1.0
	cashflow_sub_list['ProceedsFromShortTermBorrowings']		=	-1.0
	cashflow_sub_list['DividendsPaidClassifiedAsFinancingActivities']	=	-1.0
	cashflow_sub_list['CashAndCashEquivalentsAtBeginningOfPeriodCf']	=	-1.0
	cashflow_sub_list['CashAndCashEquivalentsAtEndOfPeriodCf']			=	-1.0

	unit = 100000000.0
	
	for j in range(len(raw_data)):

		word_list = raw_data[j]
		
		bool_a = (word_list[11].strip() == "당기순이익") or \
				(word_list[11].strip() == "연결당기순이익") or \
				(word_list[11].strip() == "당기순이익(손실)") or \
				(word_list[11].strip() == "연결분기순이익") 
		bool_b = (word_list[11].strip() == "배당금의 지급") or (word_list[11].strip() == "배당금지급") 
		bool_c = (word_list[11].strip() == "유형자산의 취득") or (word_list[11].strip() == "유형자산의취득")
		bool_d = (word_list[11].strip() == "무형자산의 취득") or (word_list[11].strip() == "무형자산의취득")
		bool_e = (word_list[11].strip() == "투자부동산의 취득") or (word_list[11].strip() == "투자부동산의취득")
		bool_f = (word_list[11].strip() == "유형자산의 처분") or (word_list[11].strip() == "유형자산의처분")
		bool_g = (word_list[11].strip() == "무형자산의 처분") or (word_list[11].strip() == "무형자산의처분")
		bool_h = (word_list[11].strip() == "투자부동산의 처분") or (word_list[11].strip() == "투자부동산의처분")

		#영업활동현금흐름
		if (word_list[2] == corp) and ((word_list[10] == "ifrs_CashFlowsFromUsedInOperatingActivities") or \
			((re.compile('entity[a-zA-Z0-9_]*_StatementOfCashFlowsAbstract').search(word_list[10])) and (word_list[11].strip() == "영업활동 현금흐름"))):
			if word_list[index].strip() != "":
				cashflow_sub_list['CashFlowsFromUsedInOperatingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#당기순이익(손실)
		elif (word_list[2] == corp) and \
				((word_list[10] == "dart_ProfitLossForStatementOfCashFlows") or \
				((re.compile('entity[a-zA-Z0-9_]*_StatementOfCashFlowsAbstract').search(word_list[10])) and bool_a) or \
				((re.compile('entity[a-zA-Z0-9_]*_CashFlowsFromUsedInOperatingActivities').search(word_list[10])) and bool_a)):
			if word_list[index].strip() != "":
				cashflow_sub_list['ProfitLossForStatementOfCashFlows']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#당기순이익조정을 위한 가감
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_AdjustmentsForReconcileProfitLoss"):
			if word_list[index].strip() != "":
				cashflow_sub_list['AdjustmentsForReconcileProfitLoss']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#감가상각비
		elif (word_list[2] == corp) and \
				((word_list[10] == "dart_AdjustmentsForDepreciationExpense") or \
				((re.compile('entity[a-zA-Z0-9_]*_CashFlowsFromUsedInOperatingActivities').search(word_list[10])) and (word_list[11].strip() == "감가상각비"))):
			if word_list[index].strip() != "":
				cashflow_sub_list['AdjustmentsForDepreciationExpense']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#투자활동현금흐름
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_CashFlowsFromUsedInInvestingActivities"):
			if word_list[index].strip() != "":
				cashflow_sub_list['CashFlowsFromUsedInInvestingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#유형자산의 취득
		elif (word_list[2] == corp) and \
				((word_list[10] == "ifrs_PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities") or \
				((re.compile('entity[a-zA-Z0-9_]*_CashFlowsFromUsedInInvestingActivities').search(word_list[10])) and bool_c)):
			if word_list[index].strip() != "":
				cashflow_sub_list['PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#무형자산의 취득
		elif (word_list[2] == corp) and \
				((word_list[10] == "ifrs_PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities") or \
				((re.compile('entity[a-zA-Z0-9_]*_CashFlowsFromUsedInInvestingActivities').search(word_list[10])) and bool_d)):
			if word_list[index].strip() != "":
				cashflow_sub_list['PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#투자부동산의 취득
		elif (word_list[2] == corp) and \
				((word_list[10] == "dart_PurchaseOfInvestmentProperty") or \
				((re.compile('entity[a-zA-Z0-9_]*_CashFlowsFromUsedInInvestingActivities').search(word_list[10])) and bool_e)):
			if word_list[index].strip() != "":
				cashflow_sub_list['PurchaseOfInvestmentProperty']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#유형자산의 처분
		elif (word_list[2] == corp) and \
				((word_list[10] == "ifrs_ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities") or \
				((re.compile('entity[a-zA-Z0-9_]*_CashFlowsFromUsedInInvestingActivities').search(word_list[10])) and bool_f)):
			if word_list[index].strip() != "":
				cashflow_sub_list['ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#무형자산의 처분
		elif (word_list[2] == corp) and \
				((word_list[10] == "ifrs_ProceedsFromSalesOfIntangibleAssetsClassifiedAsInvestingActivities") or \
				((re.compile('entity[a-zA-Z0-9_]*_CashFlowsFromUsedInInvestingActivities').search(word_list[10])) and bool_g)):
			if word_list[index].strip() != "":
				cashflow_sub_list['ProceedsFromSalesOfIntangibleAssetsClassifiedAsInvestingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#투자부동산의 처분
		elif (word_list[2] == corp) and \
				((word_list[10] == "dart_ProceedsFromSalesOfInvestmentProperty") or \
				((re.compile('entity[a-zA-Z0-9_]*_CashFlowsFromUsedInInvestingActivities').search(word_list[10])) and bool_h)):
			if word_list[index].strip() != "":
				cashflow_sub_list['ProceedsFromSalesOfInvestmentProperty']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#재무활동현금흐름
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_CashFlowsFromUsedInFinancingActivities"):
			if word_list[index].strip() != "":
				cashflow_sub_list['CashFlowsFromUsedInFinancingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#단기차입금의 증가
		elif (word_list[2] == corp) and (word_list[10] == "dart_ProceedsFromShortTermBorrowings"):
			if word_list[index].strip() != "":
				cashflow_sub_list['ProceedsFromShortTermBorrowings']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#배당금지급
		elif (word_list[2] == corp) and \
				((word_list[10] == "ifrs_DividendsPaidClassifiedAsFinancingActivities") or \
				((re.compile('entity[a-zA-Z0-9_]*_CashFlowsFromUsedInFinancingActivities').search(word_list[10])) and bool_b)):
			if word_list[index].strip() != "":
				cashflow_sub_list['DividendsPaidClassifiedAsFinancingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#기초현금및현금성자산
		elif (word_list[2] == corp) and (word_list[10] == "dart_CashAndCashEquivalentsAtBeginningOfPeriodCf"):
			if word_list[index].strip() != "":
				cashflow_sub_list['CashAndCashEquivalentsAtBeginningOfPeriodCf']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#기말현금및현금성자산
		elif (word_list[2] == corp) and (word_list[10] == "dart_CashAndCashEquivalentsAtEndOfPeriodCf"):
			if word_list[index].strip() != "":
				cashflow_sub_list['CashAndCashEquivalentsAtEndOfPeriodCf']	=	float(word_list[index].replace(',','').replace('\"',''))/unit

	return cashflow_sub_list

def scrape_corp_code(raw_data, corp):
	
	for j in range(len(raw_data)):
		word_list = raw_data[j]
		if (word_list[2] == corp):
			code = word_list[1]
			break

	return code.strip().replace('[','').replace(']','')

def scrape_balance_sheet(raw_data, index, corp):

	balance_sheet_sub_list = {}

	# 유동자산
	balance_sheet_sub_list['CurrentAssets']						= -1.0
	balance_sheet_sub_list['CashAndCashEquivalents']			= -1.0
	balance_sheet_sub_list['ShortTermDeposits']					= -1.0
	balance_sheet_sub_list['OtherCurrentFinancialAssets']		= -1.0
	balance_sheet_sub_list['ShortTermTradeReceivable']			= -1.0
	balance_sheet_sub_list['TradeAndOtherCurrentReceivables']	= -1.0
	balance_sheet_sub_list['ShortTermOtherReceivables']			= -1.0
	balance_sheet_sub_list['ShortTermAdvancePayments']			= -1.0
	balance_sheet_sub_list['ShortTermPrepaidExpenses']			= -1.0
	balance_sheet_sub_list['Inventories']						= -1.0
	balance_sheet_sub_list['OtherCurrentNonfinancialAssets']	= -1.0
	balance_sheet_sub_list['CurrentTaxAssets']					= -1.0
	balance_sheet_sub_list['NoncurrentAssetsOrDisposal']		= -1.0

	# 비유동자산
	balance_sheet_sub_list['NoncurrentAssets']					= -1.0
	balance_sheet_sub_list['LongTermDeposits']					= -1.0	
	balance_sheet_sub_list['OtherNoncurrentFinancialAssets']	= -1.0	
	balance_sheet_sub_list['LongTermTradeAndOther']				= -1.0
	balance_sheet_sub_list['LongTermTradeReceivablesGross']		= -1.0
	balance_sheet_sub_list['PropertyPlantAndEquipment']			= -1.0
	balance_sheet_sub_list['InvestmentProperty']				= -1.0
	balance_sheet_sub_list['GoodwillGross']						= -1.0
	balance_sheet_sub_list['IntangibleAssetsOtherThanGoodwill']	= -1.0
	balance_sheet_sub_list['InvestmentAccounted']				= -1.0
	balance_sheet_sub_list['DeferredTaxAssets']					= -1.0
	balance_sheet_sub_list['OtherNonCurrentAssets']				= -1.0
	balance_sheet_sub_list['Assets']							= -1.0

	# 유동부채
	balance_sheet_sub_list['CurrentLiabilities']				= -1.0
	balance_sheet_sub_list['TradeAndOtherCurrentPayables']		= -1.0
	balance_sheet_sub_list['ShortTermTradePayables']			= -1.0
	balance_sheet_sub_list['ShortTermOtherPayables']			= -1.0
	balance_sheet_sub_list['ShortTermAdvancesCustomers']		= -1.0
	balance_sheet_sub_list['ShortTermWithholdings']				= -1.0
	balance_sheet_sub_list['ShortTermBorrowings']				= -1.0
	balance_sheet_sub_list['CurrentPortionOfLongtermBorrowings']= -1.0
	balance_sheet_sub_list['CurrentTaxLiabilities']				= -1.0
	balance_sheet_sub_list['OtherCurrentFinancialLiabilities']	= -1.0
	balance_sheet_sub_list['CurrentProvisions']					= -1.0
	balance_sheet_sub_list['OtherCurrentLiabilities']			= -1.0
	balance_sheet_sub_list['LiabilitiesIncludedInDisposal']		= -1.0

	# 비유동부채
	balance_sheet_sub_list['NoncurrentLiabilities']				= -1.0
	balance_sheet_sub_list['LongTermTradeAndOtherNonCurrent']	= -1.0
	balance_sheet_sub_list['BondsIssued']						= -1.0
	balance_sheet_sub_list['LongTermBorrowingsGross']			= -1.0
	balance_sheet_sub_list['OtherNoncurrentFinancial']			= -1.0
	balance_sheet_sub_list['NoncurrentProvisions']				= -1.0
	balance_sheet_sub_list['PostemploymentBenefitObligations']	= -1.0
	balance_sheet_sub_list['DeferredTaxLiabilities']			= -1.0
	balance_sheet_sub_list['OtherNonCurrentLiabilities']		= -1.0

	# 자본
	balance_sheet_sub_list['Liabilities']						= -1.0
	balance_sheet_sub_list['IssuedCapital']						= -1.0
	balance_sheet_sub_list['SharePremium']						= -1.0
	balance_sheet_sub_list['RetainedEarnings']					= -1.0
	balance_sheet_sub_list['Equity']							= -1.0

	unit = 100000000.0

	for j in range(len(raw_data)):

		word_list = raw_data[j]
	
		bool_a = re.compile('^매출채권').search(word_list[11].strip())
		bool_b = (re.compile('^매입채무').search(word_list[11].strip()) or re.compile('^단기매입채무').search(word_list[11].strip()))
		bool_c = re.compile('^자[ ]*본[ ]*금').search(word_list[11].strip()) or re.compile('납입자본').search(word_list[11].strip())
		bool_d = (word_list[11].strip() == "주식발행초과금") or (word_list[11].strip() == "자본잉여금") or (word_list[11].strip() == "연결자본잉여금")
		bool_e = (word_list[11].strip() == "이익잉여금") or (word_list[11].strip() == "연결이익잉여금")
		bool_f = re.compile('^유동성장기부채').search(word_list[11].strip())
		bool_g = (word_list[11].strip() == "장기매도가능금융자산") or (word_list[11].strip() == "장기금융상품") or (word_list[11].strip() == "장기금융예치금") or (word_list[11].strip() == "금융기관예치금")
		bool_h = (word_list[11].strip() == "단기금융상품") or (word_list[11].strip() == "단기금융예치금") or (word_list[11].strip() == "금융기관예치금")
		bool_i = (word_list[11].strip() == "미수금") or (word_list[11].strip() == "단기미수금")
		bool_j = (word_list[11].strip() == "선급금") or (word_list[11].strip() == "단기선급금")
		bool_k = (word_list[11].strip() == "선급비용") or (word_list[11].strip() == "단기선급비용")
		bool_l = (word_list[11].strip() == "재고자산") or (word_list[11].strip() == "단기재고자산")

		#bool_a = (word_list[11].strip() == "매출채권") or (word_list[11].strip() == "매출채권 및 기타유동채권")
		#bool_b = (word_list[11].strip() == "매입채무") or (word_list[11].strip() == "단기매입채무") or (word_list[11].strip() == "매입채무 및 기타유동채무") 

		# 유동자산
		if (word_list[2] == corp) and (word_list[10] == "ifrs_CurrentAssets"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['CurrentAssets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 현금및현금성자산
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_CashAndCashEquivalents") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentAssets').search(word_list[10])) and (word_list[11].strip() == "현금및현금성자산"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['CashAndCashEquivalents']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 단기금융상품
		elif (word_list[2] == corp) and ((word_list[10] == "dart_ShortTermDepositsNotClassifiedAsCashEquivalents") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentAssets').search(word_list[10])) and bool_h)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermDeposits']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 기타유동금융자산
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_OtherCurrentFinancialAssets"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['OtherCurrentFinancialAssets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 매출채권
		elif (word_list[2] == corp) and \
				(((word_list[10] == "dart_ShortTermTradeReceivable") and bool_a) or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentAssets').search(word_list[10])) and bool_a)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermTradeReceivable']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 기타유동채권
		elif (word_list[2] == corp) and \
				(((word_list[10] == "ifrs_TradeAndOtherCurrentReceivables"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['TradeAndOtherCurrentReceivables']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 단기미수금
		elif (word_list[2] == corp) and ((word_list[10] == "dart_ShortTermOtherReceivables") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentAssets').search(word_list[10])) and bool_i)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermOtherReceivables']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 단기선급금
		elif (word_list[2] == corp) and ((word_list[10] == "dart_ShortTermAdvancePayments") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentAssets').search(word_list[10])) and bool_j)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermAdvancePayments']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 단기선급비용
		elif (word_list[2] == corp) and ((word_list[10] == "dart_ShortTermPrepaidExpenses") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentAssets').search(word_list[10])) and bool_k)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermPrepaidExpenses']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 재고자산
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_Inventories") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentAssets').search(word_list[10])) and bool_l)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['Inventories']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 기타유동자산
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_OtherCurrentNonfinancialAssets") or \
				(word_list[10] == "dart_OtherCurrentAssets") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentAssets').search(word_list[10])) and (word_list[11]=="기타유동자산"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['OtherCurrentNonfinancialAssets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 당기법인세자산
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_CurrentTaxAssets"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['CurrentTaxAssets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 매각예정분류자산
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_NoncurrentAssetsOrDisposalGroupsClassifiedAsHeldForSaleOrAsHeldForDistributionToOwners"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['NoncurrentAssetsOrDisposal']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#비유동자산
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_NoncurrentAssets"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['NoncurrentAssets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 장기금융상품
		elif (word_list[2] == corp) and ((word_list[10] == "dart_LongTermDepositsNotClassifiedAsCashEquivalents") or \
			((re.compile('entity[a-zA-Z0-9_]*_NoncurrentAssets').search(word_list[10])) and bool_g)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['LongTermDeposits']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 기타비유동금융자산
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_OtherNoncurrentFinancialAssets"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['OtherNoncurrentFinancialAssets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 장기매출채권 및 기타비유동채권
		elif (word_list[2] == corp) and (word_list[10] == "dart_LongTermTradeAndOtherNonCurrentReceivablesGross"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['LongTermTradeAndOther']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 장기매출채권
		elif (word_list[2] == corp) and (word_list[10] == "dart_LongTermTradeReceivablesGross"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['LongTermTradeReceivablesGross']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 유형자산
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_PropertyPlantAndEquipment") or \
				((re.compile('entity[a-zA-Z0-9_]*_NoncurrentAssets').search(word_list[10])) and (word_list[11].strip() == "유형자산"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['PropertyPlantAndEquipment']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 투자부동산
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_InvestmentProperty") or \
				((re.compile('entity[a-zA-Z0-9_]*_NoncurrentAssets').search(word_list[10])) and (word_list[11].strip() == "투자부동산"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['InvestmentProperty']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 영업권
		elif (word_list[2] == corp) and ((word_list[10] == "dart_GoodwillGross") or \
				((re.compile('entity[a-zA-Z0-9_]*_NoncurrentAssets').search(word_list[10])) and (word_list[11].strip() == "영업권"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['GoodwillGross']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 영업권 이외의 무형자산
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_IntangibleAssetsOtherThanGoodwill") or \
				((re.compile('entity[a-zA-Z0-9_]*_NoncurrentAssets').search(word_list[10])) and (word_list[11].strip() == "무형자산"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['IntangibleAssetsOtherThanGoodwill']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 지분법적용 투자지분
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_InvestmentAccountedForUsingEquityMethod") or \
				((re.compile('entity[a-zA-Z0-9_]*_NoncurrentAssets').search(word_list[10])) and (word_list[11].strip() == "관계기업 및 공동기업 투자"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['InvestmentAccounted']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 이연법인세자산
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_DeferredTaxAssets"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['DeferredTaxAssets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 기타비유동자산 
		elif (word_list[2] == corp) and ((word_list[10] == "dart_OtherNonCurrentAssets") or \
				((re.compile('entity[a-zA-Z0-9_]*_NoncurrentAssets').search(word_list[10])) and (word_list[11].strip() == "기타비유동자산"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['OtherNonCurrentAssets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#자산총계
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_Assets"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['Assets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#유동부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_CurrentLiabilities"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['CurrentLiabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#매입채무 및 기타유동부채
		elif (word_list[2] == corp) and \
				(((word_list[10] == "ifrs_TradeAndOtherCurrentPayables") and bool_b) or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentLiabilities').search(word_list[10])) and bool_b)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['TradeAndOtherCurrentPayables']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#단기매입채무
		elif (word_list[2] == corp) and (word_list[10] == "dart_ShortTermTradePayables"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermTradePayables']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#단기미지급금
		elif (word_list[2] == corp) and ((word_list[10] == "dart_ShortTermOtherPayables") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentLiabilities').search(word_list[10])) and (word_list[11].strip() == "미지급금"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermOtherPayables']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
      	#단기선수금
		elif (word_list[2] == corp) and ((word_list[10] == "dart_ShortTermAdvancesCustomers") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentLiabilities').search(word_list[10])) and (word_list[11].strip() == "선수금"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermAdvancesCustomers']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
      	#단기예수금
		elif (word_list[2] == corp) and ((word_list[10] == "dart_ShortTermWithholdings") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentLiabilities').search(word_list[10])) and (word_list[11].strip() == "예수금"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermWithholdings']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#단기차입금
		elif (word_list[2] == corp) and ((word_list[10] == "dart_ShortTermBorrowings") or (word_list[10] == "ifrs_ShorttermBorrowings")):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermBorrowings']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#유동성장기차입금
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_CurrentPortionOfLongtermBorrowings") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentLiabilities').search(word_list[10])) and bool_f)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['CurrentPortionOfLongtermBorrowings']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
      	#당기법인세부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_CurrentTaxLiabilities"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['CurrentTaxLiabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
      	#기타유동금융부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_OtherCurrentFinancialLiabilities"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['OtherCurrentFinancialLiabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
      	#유동충당부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_CurrentProvisions"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['CurrentProvisions']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
      	#기타유동부채
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_OtherCurrentNonfinancialLiabilities") or \
				(word_list[10] == "dart_OtherCurrentLiabilities") or \
				((re.compile('entity[a-zA-Z0-9_]*_CurrentLiabilities').search(word_list[10])) and (word_list[11].strip() == "기타유동부채"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['OtherCurrentLiabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
      	#매각예정분류부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_LiabilitiesIncludedInDisposalGroupsClassifiedAsHeldForSale"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['LiabilitiesIncludedInDisposal']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#비유동부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_NoncurrentLiabilities"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['NoncurrentLiabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 장기매입채무 및 기타비유동채무
		elif (word_list[2] == corp) and (word_list[10] == "dart_LongTermTradeAndOtherNonCurrentPayables"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['LongTermTradeAndOtherNonCurrent']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 사채
		elif (word_list[2] == corp) and ((word_list[10] == "dart_BondsIssued") or \
				((re.compile('entity[a-zA-Z0-9_]*_NoncurrentLiabilities').search(word_list[10])) and (word_list[11].strip() == "사채"))):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['BondsIssued']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#장기차입금
		elif (word_list[2] == corp) and (word_list[10] == "dart_LongTermBorrowingsGross"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['LongTermBorrowingsGross']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 기타비유동금융부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_OtherNoncurrentFinancialLiabilities"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['OtherNoncurrentFinancial']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 비유동충당부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_NoncurrentProvisions"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['NoncurrentProvisions']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 퇴직급여부채
		elif (word_list[2] == corp) and (word_list[10] == "dart_PostemploymentBenefitObligations"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['PostemploymentBenefitObligations']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 이연법인세부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_DeferredTaxLiabilities"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['DeferredTaxLiabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 기타비유동부채
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_OtherNoncurrentNonfinancialLiabilities") or \
				(word_list[10] == "dart_OtherNonCurrentLiabilities")):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['OtherNonCurrentLiabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#부채총계
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_Liabilities"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['Liabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#자본금
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_IssuedCapital") or \
				(word_list[10] == "dart_ContributedEquity") or \
				((re.compile('entity[a-zA-Z0-9_]*_EquityAttributableToOwnersOfParent').search(word_list[10])) and bool_c) or \
				((re.compile('entity[a-zA-Z0-9_]*_EquityAbstract').search(word_list[10])) and bool_c)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['IssuedCapital']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 자본잉여금
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_SharePremium") or \
				((word_list[10] == "dart_CapitalSurplus") and (word_list[11].strip() == "자본잉여금")) or \
				((word_list[10] == "dart_ElementsOfOtherStockholdersEquity") and (word_list[11].strip() == "자본잉여금")) or \
				((re.compile('entity[a-zA-Z0-9_]*_EquityAttributableToOwnersOfParent').search(word_list[10])) and bool_d) or \
				((re.compile('entity[a-zA-Z0-9_]*_ContributedEquity').search(word_list[10])) and bool_d) or \
				((re.compile('entity[a-zA-Z0-9_]*_EquityAbstract').search(word_list[10])) and bool_d)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['SharePremium']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#이익잉여금(결손금)
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_RetainedEarnings") or \
				((re.compile('entity[a-zA-Z0-9_]*_EquityAttributableToOwnersOfParent').search(word_list[10])) and bool_e) or \
				((re.compile('entity[a-zA-Z0-9_]*_ContributedEquity').search(word_list[10])) and bool_e) or \
				((re.compile('entity[a-zA-Z0-9_]*_EquityAbstract').search(word_list[10])) and bool_e)):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['RetainedEarnings']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#자본총계
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_Equity"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['Equity']	=	float(word_list[index].replace(',','').replace('\"',''))/unit

	return balance_sheet_sub_list

def scrape_income_statement(raw_data, index, corp):

	revenue				= -1.0
	op_income			= -1.0
	profit				= -1.0
	cost_of_sale		= -1.0
	gross_profit		= -1.0
	admin_expenses		= -1.0
	profit_before_tax	= -1.0
	income_tax			= -1.0
	basic_eps			= -1.0
	other_gain			= -1.0
	other_loss			= -1.0
	finance_income		= -1.0
	finance_cost		= -1.0
	associates_profit	= -1.0

	unit = 100000000.0
	
	for j in range(len(raw_data)):

		word_list = raw_data[j]

		bool_a = (word_list[11].strip() == "지분법이익") or (word_list[11].strip() == "공동기업및관계기업투자손익") or (word_list[11].strip() == "지분법투자관련 손익") or (word_list[11].strip() == "지분법관련손익") or (word_list[11].strip() == "종속기업및관계기업투자관련이익")
		bool_b = re.compile('보통주').search(word_list[11].strip()) or re.compile('^기본주당순이익').search(word_list[11].strip())
		bool_c = (word_list[11].strip() == "기타수익") or (word_list[11].strip() == "기타영업외수익") 
		bool_d = (word_list[11].strip() == "기타비용") or (word_list[11].strip() == "기타영업외비용") 
		bool_e = (word_list[11].strip() == "당기순이익") or (word_list[11].strip() == "연결당기순이익") 
		bool_f = (word_list[11].strip() == "법인세차감전계속영업순이익") or (word_list[11].strip() == "법인세비용차감전계속영업순이익(손실)")
		bool_g = (word_list[11].strip() == "계속영업법인세비용") or (word_list[11].strip() == "계속영업손익법인세비용(효익)")

		# 매출액 or 영업수익
		if (word_list[2] == corp) and (word_list[10] == "ifrs_Revenue"):
			if word_list[index].strip() != "":
				revenue = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 매출원가
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_CostOfSales") or \
			(word_list[10] == "dart_OperatingExpenses") or \
			((re.compile('entity[a-zA-Z0-9_]*_IncomeStatementAbstract').search(word_list[10])) and (word_list[11].strip() == "영업비용"))):
			if word_list[index].strip() != "":
				cost_of_sale = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 매출총이익
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_GrossProfit")):
			if word_list[index].strip() != "":
				gross_profit = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 판매비와 관리비
		elif (word_list[2] == corp) and (word_list[10] == "dart_TotalSellingGeneralAdministrativeExpenses"):
			if word_list[index].strip() != "":
				admin_expenses = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 영업이익
		elif (word_list[2] == corp) and ((word_list[10] == "dart_OperatingIncomeLoss") or \
			((re.compile('entity[a-zA-Z0-9_]*_IncomeStatementAbstract').search(word_list[10])) and (word_list[11].strip() == "영업이익"))):
			if word_list[index].strip() != "":
				op_income = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 법인세비용차감전순이익
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_ProfitLossBeforeTax") or \
			((re.compile('entity[a-zA-Z0-9_]*_StatementOfComprehensiveIncomeAbstract').search(word_list[10])) and bool_f) or \
			((re.compile('entity[a-zA-Z0-9_]*_IncomeStatementAbstract').search(word_list[10])) and bool_f)):
			if word_list[index].strip() != "":
				profit_before_tax = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 법인세비용
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_IncomeTaxExpenseContinuingOperations") or \
			((re.compile('entity[a-zA-Z0-9_]*_ProfitLossBeforeTax').search(word_list[10])) and bool_g) or \
			((re.compile('entity[a-zA-Z0-9_]*_StatementOfComprehensiveIncomeAbstract').search(word_list[10])) and bool_g) or \
			((re.compile('entity[a-zA-Z0-9_]*_IncomeStatementAbstract').search(word_list[10])) and bool_g)):
			if word_list[index].strip() != "":
				income_tax = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 당기순이익
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_ProfitLoss") or \
				((re.compile('entity[a-zA-Z0-9_]*_IncomeStatementAbstract').search(word_list[10])) and bool_e)):
			if word_list[index].strip() != "":
				profit = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 기본주당순이익
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_BasicEarningsLossPerShare") or \
				(word_list[10] == "ifrs_BasicEarningsLossPerShare") or \
				((re.compile('entity[a-zA-Z0-9_]*_IncomeStatementAbstract').search(word_list[10])) and bool_b) or \
				((re.compile('entity[a-zA-Z0-9_]*_BasicEarningsLossPerShare').search(word_list[10])) and bool_b)):
			if word_list[index].strip() != "":
				basic_eps = float(word_list[index].replace(",","").replace('\"',''))
		#기타수익
		elif (word_list[2] == corp) and ((word_list[10] == "dart_OtherGains") or \
				((re.compile('entity[a-zA-Z0-9_]*_OperatingIncomeLoss').search(word_list[10])) and bool_c) or \
				((re.compile('entity[a-zA-Z0-9_]*_IncomeStatementAbstract').search(word_list[10])) and bool_c)):
			if word_list[index].strip() != "":
				other_gain = float(word_list[index].replace(",","").replace('\"',''))/unit
		#기타비용
		elif (word_list[2] == corp) and ((word_list[10] == "dart_OtherLosses") or \
				((re.compile('entity[a-zA-Z0-9_]*_OperatingIncomeLoss').search(word_list[10])) and bool_d) or \
				((re.compile('entity[a-zA-Z0-9_]*_IncomeStatementAbstract').search(word_list[10])) and bool_d)):
			if word_list[index].strip() != "":
				other_loss = float(word_list[index].replace(",","").replace('\"',''))/unit
		#금융수익
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_FinanceIncome") or \
				((re.compile('entity[a-zA-Z0-9_]*_OperatingIncomeLoss').search(word_list[10])) and (word_list[11].strip() == "금융수익")) or \
				((re.compile('entity[a-zA-Z0-9_]*_IncomeStatementAbstract').search(word_list[10])) and (word_list[11].strip() == "금융수익"))):
			if word_list[index].strip() != "":
				finance_income = float(word_list[index].replace(",","").replace('\"',''))/unit
		#금융비용
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_FinanceCosts") or \
				((re.compile('entity[a-zA-Z0-9_]*_OperatingIncomeLoss').search(word_list[10])) and (word_list[11].strip() == "금융비용")) or \
				((re.compile('entity[a-zA-Z0-9_]*_IncomeStatementAbstract').search(word_list[10])) and (word_list[11].strip() == "금융비용"))):
			if word_list[index].strip() != "":
				finance_cost = float(word_list[index].replace(",","").replace('\"',''))/unit
		#지분법이익
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_ShareOfProfitLossOfAssociatesAndJointVenturesAccountedForUsingEquityMethod") or \
				((re.compile('entity[a-zA-Z0-9_]*_StatementOfComprehensiveIncomeAbstract').search(word_list[10])) and bool_a) or \
				((re.compile('entity[a-zA-Z0-9_]*_IncomeStatementAbstract').search(word_list[10])) and bool_a)):
			if word_list[index].strip() != "":
				associates_profit = float(word_list[index].replace(",","").replace('\"',''))/unit

	#print(revenue)
	#print(op_income)
	#print(profit)
	#print(cost_of_sale)
	#print(gross_profit)
	#print(admin_expenses)
	#print(profit_before_tax)
	#print(income_tax)
	#print(basic_eps)
	
	return [revenue, cost_of_sale, gross_profit, admin_expenses, op_income, other_gain, other_loss, finance_income, finance_cost, associates_profit, profit_before_tax, income_tax, profit, basic_eps]

def get_cashflow_statement_year(corp):

	cashflow_statement_sub_list = []

	#2015 4Q
	file_name = "2015.4Q/2015_사업보고서_04_현금흐름표_연결_20160601.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 15, corp))
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 14, corp))
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2016 4Q
	file_name = "2016.4Q/2016_사업보고서_04_현금흐름표_연결_20170524.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))

	#2017 4Q
	#file_name = "2017.3Q/2017_3분기보고서_04_현금흐름표_연결_20180131.txt"
	file_name = "2017.4Q/2017_사업보고서_04_현금흐름표_연결_20180521.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))

	return cashflow_statement_sub_list
	
def get_individual_cashflow_statement_year(corp):

	cashflow_statement_sub_list = []

	#2015 4Q
	file_name = "2015.4Q/2015_사업보고서_04_현금흐름표_20160601.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 15, corp))
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 14, corp))
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2016 4Q
	file_name = "2016.4Q/2016_사업보고서_04_현금흐름표_20170524.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))

	#2017 4Q
	#file_name = "2017.3Q/2017_3분기보고서_04_현금흐름표_20180131.txt"
	file_name = "2017.4Q/2017_사업보고서_04_현금흐름표_20180521.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))

	return cashflow_statement_sub_list
	
def get_cashflow_statement_quarter(corp):

	cashflow_statement_sub_list = []

	#2016 1Q
	file_name = "2016.1Q/2016_1분기보고서_04_현금흐름표_연결_20160818.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2016 2Q
	file_name = "2016.2Q/2016_반기보고서_04_현금흐름표_연결_20161115.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2016 3Q
	file_name = "2016.3Q/2016_3분기보고서_04_현금흐름표_연결_20170223.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2016 4Q
	file_name = "2016.4Q/2016_사업보고서_04_현금흐름표_연결_20170524.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))

	#2017 1Q
	file_name = "2017.1Q/2017_1분기보고서_04_현금흐름표_연결_20170816.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2017 2Q
	file_name = "2017.2Q/2017_반기보고서_04_현금흐름표_연결_20171102.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2017 3Q
	file_name = "2017.3Q/2017_3분기보고서_04_현금흐름표_연결_20180131.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2017 4Q
	file_name = "2017.4Q/2017_사업보고서_04_현금흐름표_연결_20180521.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))

	return cashflow_statement_sub_list

def get_individual_cashflow_statement_quarter(corp):

	cashflow_statement_sub_list = []

	#2016 1Q
	file_name = "2016.1Q/2016_1분기보고서_04_현금흐름표_20160818.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2016 2Q
	file_name = "2016.2Q/2016_반기보고서_04_현금흐름표_20161115.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2016 3Q
	file_name = "2016.3Q/2016_3분기보고서_04_현금흐름표_20170223.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2016 4Q
	file_name = "2016.4Q/2016_사업보고서_04_현금흐름표_20170524.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))

	#2017 1Q
	file_name = "2017.1Q/2017_1분기보고서_04_현금흐름표_20170816.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2017 2Q
	file_name = "2017.2Q/2017_반기보고서_04_현금흐름표_20171102.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2017 3Q
	file_name = "2017.3Q/2017_3분기보고서_04_현금흐름표_20180131.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))
	
	#2017 4Q
	file_name = "2017.4Q/2017_사업보고서_04_현금흐름표_20180521.txt"
	raw_data = read_raw_data(file_name)
	cashflow_statement_sub_list.append(scrape_cashflow_statement(raw_data, 12, corp))

	return cashflow_statement_sub_list

def get_balance_sheet_year(corp):

	balance_sheet_sub_list = []

	#2015 4Q
	file_name = "2015.4Q/2015_사업보고서_01_재무상태표_연결_20160531.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 14, corp))
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 13, corp))
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2016 4Q
	file_name = "2016.4Q/2016_사업보고서_01_재무상태표_연결_20170524.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))

	#2017 4Q
	#file_name = "2017.3Q/2017_3분기보고서_01_재무상태표_연결_20180131.txt"
	file_name = "2017.4Q/2017_사업보고서_01_재무상태표_연결_20180521.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))

	return balance_sheet_sub_list
	
def get_balance_sheet_quarter(corp):

	balance_sheet_sub_list = []

	#2016 1Q
	file_name = "2016.1Q/2016_1분기보고서_01_재무상태표_연결_20160818.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2016 2Q
	file_name = "2016.2Q/2016_반기보고서_01_재무상태표_연결_20161115.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2016 3Q
	file_name = "2016.3Q/2016_3분기보고서_01_재무상태표_연결_20170223.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2016 4Q
	file_name = "2016.4Q/2016_사업보고서_01_재무상태표_연결_20170524.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))

	#2017 1Q
	file_name = "2017.1Q/2017_1분기보고서_01_재무상태표_연결_20170816.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2017 2Q
	file_name = "2017.2Q/2017_반기보고서_01_재무상태표_연결_20171102.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2017 3Q
	file_name = "2017.3Q/2017_3분기보고서_01_재무상태표_연결_20180131.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2017 4Q
	file_name = "2017.4Q/2017_사업보고서_01_재무상태표_연결_20180521.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))

	return balance_sheet_sub_list

def get_individual_balance_sheet_year(corp):

	balance_sheet_sub_list = []

	#2015 4Q
	file_name = "2015.4Q/2015_사업보고서_01_재무상태표_20160531.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 14, corp))
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 13, corp))
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2016 4Q
	file_name = "2016.4Q/2016_사업보고서_01_재무상태표_20170524.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))

	#2017 4Q
	#file_name = "2017.3Q/2017_3분기보고서_01_재무상태표_20180131.txt"
	file_name = "2017.4Q/2017_사업보고서_01_재무상태표_20180521.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))

	code = scrape_corp_code(raw_data, corp)
	#code = "175330"
	
	return balance_sheet_sub_list, code
	
def get_individual_balance_sheet_quarter(corp):

	balance_sheet_sub_list = []

	#2016 1Q
	file_name = "2016.1Q/2016_1분기보고서_01_재무상태표_20160818.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2016 2Q
	file_name = "2016.2Q/2016_반기보고서_01_재무상태표_20161115.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2016 3Q
	file_name = "2016.3Q/2016_3분기보고서_01_재무상태표_20170223.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2016 4Q
	file_name = "2016.4Q/2016_사업보고서_01_재무상태표_20170524.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))

	#2017 1Q
	file_name = "2017.1Q/2017_1분기보고서_01_재무상태표_20170816.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2017 2Q
	file_name = "2017.2Q/2017_반기보고서_01_재무상태표_20171102.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2017 3Q
	file_name = "2017.3Q/2017_3분기보고서_01_재무상태표_20180131.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))
	
	#2017 4Q
	file_name = "2017.4Q/2017_사업보고서_01_재무상태표_20180521.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))

	return balance_sheet_sub_list

def get_income_statement_year(corp, mode):

	income_statement_sub_list = []

	#2015 4Q
	if mode == 0:
		file_name = "2015.4Q/2015_사업보고서_02_손익계산서_연결_20160531.txt"
	else:
		file_name = "2015.4Q/2015_사업보고서_03_포괄손익계산서_연결_20160531.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 17, corp))
	income_statement_sub_list.append(scrape_income_statement(raw_data, 16, corp))
	income_statement_sub_list.append(scrape_income_statement(raw_data, 13, corp))
	
	#2016 4Q
	if mode == 0:
		file_name = "2016.4Q/2016_사업보고서_02_손익계산서_연결_20170524.txt"
	else:
		file_name = "2016.4Q/2016_사업보고서_03_포괄손익계산서_연결_20170524.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 13, corp))

	#2017 4Q
	if mode == 0:
		#file_name = "2017.3Q/2017_3분기보고서_02_손익계산서_연결_20180131.txt"
		file_name = "2017.4Q/2017_사업보고서_02_손익계산서_연결_20180521.txt"
	else:
		file_name = "2017.4Q/2017_사업보고서_03_포괄손익계산서_연결_20180521.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 13, corp))

	return income_statement_sub_list
	
def get_individual_income_statement_year(corp, mode):

	income_statement_sub_list = []

	#2015 4Q
	if mode == 0:
		file_name = "2015.4Q/2015_사업보고서_02_손익계산서_20160531.txt"
	else:
		file_name = "2015.4Q/2015_사업보고서_03_포괄손익계산서_20160531.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 17, corp))
	income_statement_sub_list.append(scrape_income_statement(raw_data, 16, corp))
	income_statement_sub_list.append(scrape_income_statement(raw_data, 13, corp))
	
	#2016 4Q
	if mode == 0:
		file_name = "2016.4Q/2016_사업보고서_02_손익계산서_20170524.txt"
	else:
		file_name = "2016.4Q/2016_사업보고서_03_포괄손익계산서_20170524.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 13, corp))

	#2017 4Q
	if mode == 0:
		#file_name = "2017.3Q/2017_3분기보고서_02_손익계산서_20180131.txt"
		file_name = "2017.4Q/2017_사업보고서_02_손익계산서_20180521.txt"
	else:
		#file_name = "2017.3Q/2017_3분기보고서_03_포괄손익계산서_20180131.txt"
		file_name = "2017.4Q/2017_사업보고서_03_포괄손익계산서_20180521.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 13, corp))

	return income_statement_sub_list
	
def get_income_statement_quarter(corp, mode):

	income_statement_sub_list = []

	#2016 1Q
	if mode == 0:
		file_name = "2016.1Q/2016_1분기보고서_02_손익계산서_연결_20160818.txt"
	else:
		file_name = "2016.1Q/2016_1분기보고서_03_포괄손익계산서_연결_20160818.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2016 2Q
	if mode == 0:
		file_name = "2016.2Q/2016_반기보고서_02_손익계산서_연결_20161115.txt"
	else:
		file_name = "2016.2Q/2016_반기보고서_03_포괄손익계산서_연결_20161115.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2016 3Q
	if mode == 0:
		file_name = "2016.3Q/2016_3분기보고서_02_손익계산서_연결_20170223.txt"
	else:
		file_name = "2016.3Q/2016_3분기보고서_03_포괄손익계산서_연결_20170223.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2016 4Q
	if mode == 0:
		file_name = "2016.4Q/2016_사업보고서_02_손익계산서_연결_20170524.txt"
	else:
		file_name = "2016.4Q/2016_사업보고서_03_포괄손익계산서_연결_20170524.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 13, corp))
	for i in range(len(income_statement_sub_list[3])):
		if (income_statement_sub_list[3][i] != -1) and (income_statement_sub_list[2][i] != -1) and (income_statement_sub_list[1][i] != -1) and (income_statement_sub_list[0][i] != -1):
			income_statement_sub_list[3][i] = income_statement_sub_list[3][i] - income_statement_sub_list[2][i] - income_statement_sub_list[1][i] - income_statement_sub_list[0][i]
		else:
			income_statement_sub_list[3][i] = -1

	#2017 1Q
	if mode == 0:
		file_name = "2017.1Q/2017_1분기보고서_02_손익계산서_연결_20170816.txt"
	else:
		file_name = "2017.1Q/2017_1분기보고서_03_포괄손익계산서_연결_20170816.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2017 2Q
	if mode == 0:
		file_name = "2017.2Q/2017_반기보고서_02_손익계산서_연결_20171102.txt"
	else:
		file_name = "2017.2Q/2017_반기보고서_03_포괄손익계산서_연결_20171102.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2017 3Q
	if mode == 0:
		file_name = "2017.3Q/2017_3분기보고서_02_손익계산서_연결_20180131.txt"
	else:
		file_name = "2017.3Q/2017_3분기보고서_03_포괄손익계산서_연결_20180131.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2017 4Q
	if mode == 0:
		file_name = "2017.4Q/2017_사업보고서_02_손익계산서_연결_20180521.txt"
	else:
		file_name = "2017.4Q/2017_사업보고서_03_포괄손익계산서_연결_20180521.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 13, corp))
	for i in range(len(income_statement_sub_list[7])):
		if (income_statement_sub_list[7][i] != -1) and (income_statement_sub_list[6][i] != -1) and (income_statement_sub_list[5][i] != -1) and (income_statement_sub_list[4][i] != -1):
			income_statement_sub_list[7][i] = income_statement_sub_list[7][i] - income_statement_sub_list[6][i] - income_statement_sub_list[5][i] - income_statement_sub_list[4][i]
		else:
			income_statement_sub_list[7][i] = -1

	return income_statement_sub_list

def get_individual_income_statement_quarter(corp, mode):

	income_statement_sub_list = []

	#2016 1Q
	if mode == 0:
		file_name = "2016.1Q/2016_1분기보고서_02_손익계산서_20160818.txt"
	else:
		file_name = "2016.1Q/2016_1분기보고서_03_포괄손익계산서_20160818.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2016 2Q
	if mode == 0:
		file_name = "2016.2Q/2016_반기보고서_02_손익계산서_20161115.txt"
	else:
		file_name = "2016.2Q/2016_반기보고서_03_포괄손익계산서_20161115.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2016 3Q
	if mode == 0:
		file_name = "2016.3Q/2016_3분기보고서_02_손익계산서_20170223.txt"
	else:
		file_name = "2016.3Q/2016_3분기보고서_03_포괄손익계산서_20170223.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2016 4Q
	if mode == 0:
		file_name = "2016.4Q/2016_사업보고서_02_손익계산서_20170524.txt"
	else:
		file_name = "2016.4Q/2016_사업보고서_03_포괄손익계산서_20170524.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 13, corp))
	for i in range(len(income_statement_sub_list[3])):
		if (income_statement_sub_list[3][i] != -1) and (income_statement_sub_list[2][i] != -1) and (income_statement_sub_list[1][i] != -1) and (income_statement_sub_list[0][i] != -1):
			income_statement_sub_list[3][i] = income_statement_sub_list[3][i] - income_statement_sub_list[2][i] - income_statement_sub_list[1][i] - income_statement_sub_list[0][i]
		else:
			income_statement_sub_list[3][i] = -1
	#2017 1Q
	if mode == 0:
		file_name = "2017.1Q/2017_1분기보고서_02_손익계산서_20170816.txt"
	else:
		file_name = "2017.1Q/2017_1분기보고서_03_포괄손익계산서_20170816.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2017 2Q
	if mode == 0:
		file_name = "2017.2Q/2017_반기보고서_02_손익계산서_20171102.txt"
	else:
		file_name = "2017.2Q/2017_반기보고서_03_포괄손익계산서_20171102.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2017 3Q
	if mode == 0:
		file_name = "2017.3Q/2017_3분기보고서_02_손익계산서_20180131.txt"
	else:
		file_name = "2017.3Q/2017_3분기보고서_03_포괄손익계산서_20180131.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 12, corp))
	
	#2017 4Q
	if mode == 0:
		file_name = "2017.4Q/2017_사업보고서_02_손익계산서_20180521.txt"
	else:
		file_name = "2017.4Q/2017_사업보고서_03_포괄손익계산서_20180521.txt"
	raw_data = read_raw_data(file_name)
	income_statement_sub_list.append(scrape_income_statement(raw_data, 13, corp))
	for i in range(len(income_statement_sub_list[7])):
		if (income_statement_sub_list[7][i] != -1) and (income_statement_sub_list[6][i] != -1) and (income_statement_sub_list[5][i] != -1) and (income_statement_sub_list[4][i] != -1):
			income_statement_sub_list[7][i] = income_statement_sub_list[7][i] - income_statement_sub_list[6][i] - income_statement_sub_list[5][i] - income_statement_sub_list[4][i]
		else:
			income_statement_sub_list[7][i] = -1
	return income_statement_sub_list

def unzip(source_file, dest_path):
	
	with zipfile.ZipFile(source_file, 'r') as zf:
		zipInfo = zf.infolist()
		
		for member in zipInfo:
			print(type(member.filename))
			#member.filename = member.filename.decode("euc-kr").encode("utf-8")
			member.filename = bytes(member.filename, encoding="cp437").decode("euc-kr")
			zf.extract(member, dest_path)

def zip_test():
	print("zip test")
	url_list = ["download_test/2015_4Q_BS_20160531132458.zip", "download_test/2015_4Q_PL_20160531132719.zip", "download_test/2015_4Q_CF_20160601132810.zip", "download_test/2015_4Q_CE_20160531133335.zip"]
	dest_dir = "download_test/2015.4Q"
	for i in range(len(url_list)):
		#handle=urllib.request.urlretrieve("http://dart.fss.or.kr/api/xbrl/"+url_list[i], url_list[i])
		unzip(url_list[i], dest_dir)
	
def download_files():

	# 2015
	url_list = ["2015_4Q_BS_20160531132458.zip", "2015_4Q_PL_20160531132719.zip", "2015_4Q_CF_20160601132810.zip", "2015_4Q_CE_20160531133335.zip"]
	dest_dir = "2015.4Q"
	for i in range(len(url_list)):
		handle=urllib.request.urlretrieve("http://dart.fss.or.kr/api/xbrl/download/"+url_list[i], url_list[i])
		unzip(url_list[i], dest_dir)
		os.remove(url_list[i])

	#2016 1Q
	url_list = ["2016_1Q_BS_20160818183809.zip", "2016_1Q_PL_20160818183926.zip", "2016_1Q_CF_20160818184055.zip","2016_1Q_CE_20160818184320.zip"]
	dest_dir = "2016.1Q"
	for i in range(len(url_list)):
		handle=urllib.request.urlretrieve("http://dart.fss.or.kr/api/xbrl/download/"+url_list[i], url_list[i])
		unzip(url_list[i], dest_dir)
		os.remove(url_list[i])
	
	#2016 2Q
	url_list = ["2016_2Q_BS_20161115192253.zip", "2016_2Q_PL_20161115192420.zip", "2016_2Q_CF_20161115192602.zip","2016_2Q_CE_20161115192840.zip"]
	dest_dir = "2016.2Q"
	for i in range(len(url_list)):
		handle=urllib.request.urlretrieve("http://dart.fss.or.kr/api/xbrl/download/"+url_list[i], url_list[i])
		unzip(url_list[i], dest_dir)
		os.remove(url_list[i])
	
	#2016 3Q
	url_list = ["2016_3Q_BS_20170223102846.zip", "2016_3Q_PL_20170223103051.zip", "2016_3Q_CF_20170223103257.zip", "2016_3Q_CE_20170223103604.zip"]
	dest_dir = "2016.3Q"
	for i in range(len(url_list)):
		handle=urllib.request.urlretrieve("http://dart.fss.or.kr/api/xbrl/download/"+url_list[i], url_list[i])
		unzip(url_list[i], dest_dir)
		os.remove(url_list[i])
	
	#2016 4Q
	url_list = ["2016_4Q_BS_20170524161103.zip", "2016_4Q_PL_20170524161337.zip", "2016_4Q_CF_20170524161503.zip","2016_4Q_CE_20170524161706.zip"]
	dest_dir = "2016.4Q"
	for i in range(len(url_list)):
		handle=urllib.request.urlretrieve("http://dart.fss.or.kr/api/xbrl/download/"+url_list[i], url_list[i])
		unzip(url_list[i], dest_dir)
		os.remove(url_list[i])
	
	#2017 1Q                                                                    
	url_list = ["2017_1Q_BS_20170816150254.zip", "2017_1Q_PL_20170816150736.zip", "2017_1Q_CF_20170816151052.zip", "2017_1Q_CE_20170816151506.zip"]
	dest_dir = "2017.1Q"
	for i in range(len(url_list)):
		handle=urllib.request.urlretrieve("http://dart.fss.or.kr/api/xbrl/download/"+url_list[i], url_list[i])
		unzip(url_list[i], dest_dir)
		os.remove(url_list[i])
	
	#2017 2Q                                                                    
	url_list = ["2017_2Q_BS_20171102155644.zip", "2017_2Q_PL_20171102160044.zip", "2017_2Q_CF_20171102160315.zip", "2017_2Q_CE_20171102160615.zip"]
	dest_dir = "2017.2Q"
	for i in range(len(url_list)):
		handle=urllib.request.urlretrieve("http://dart.fss.or.kr/api/xbrl/download/"+url_list[i], url_list[i])
		unzip(url_list[i], dest_dir)
		os.remove(url_list[i])
	
	#2017 3Q                                                                    
	url_list = ["2017_3Q_BS_20180131152406.zip", "2017_3Q_PL_20180131152743.zip", "2017_3Q_CF_20180131153003.zip", "2017_3Q_CE_20180131153246.zip"]
	dest_dir = "2017.3Q"
	for i in range(len(url_list)):
		handle=urllib.request.urlretrieve("http://dart.fss.or.kr/api/xbrl/download/"+url_list[i], url_list[i])
		unzip(url_list[i], dest_dir)
		os.remove(url_list[i])

	#2017 4Q                                                                    
	url_list = ["2017_4Q_BS_20180521114003.zip", "2017_4Q_PL_20180521114402.zip", "2017_4Q_CF_20180521114637.zip", "2017_4Q_CE_20180521114935.zip"]
	dest_dir = "2017.4Q"
	for i in range(len(url_list)):
		handle=urllib.request.urlretrieve("http://dart.fss.or.kr/api/xbrl/download/"+url_list[i], url_list[i])
		unzip(url_list[i], dest_dir)
		os.remove(url_list[i])
	
#form_class = uic.loadUiType("main_window.ui")[0]
#
#class MyWindow(QMainWindow, form_class):
#    def __init__(self):
#        super().__init__()
#        self.setupUi(self)

def run_dart(corp):
	# Read income statement
	print("Read income statement")
	income_statement_sub_list	= get_income_statement_year(corp, 0)
	income_statement_sub_list_q = get_income_statement_quarter(corp, 0)

	# Read individual income statement
	print("Read individual income statement")
	individual_income_statement_sub_list	= get_individual_income_statement_year(corp, 0)
	individual_income_statement_sub_list_q	= get_individual_income_statement_quarter(corp, 0)

	income_statement_sub_list.reverse()
	income_statement_sub_list_q.reverse()
	individual_income_statement_sub_list.reverse()
	individual_income_statement_sub_list_q.reverse()

	# Read comprehensive income statement
	print("Read comprehensive income statement")
	comp_income_statement_sub_list	= get_income_statement_year(corp, 1)
	comp_income_statement_sub_list_q = get_income_statement_quarter(corp, 1)

	# Read individual comprehensive income statement
	print("Read individual comprehensive income statement")
	individual_comp_income_statement_sub_list	= get_individual_income_statement_year(corp, 1)
	individual_comp_income_statement_sub_list_q	= get_individual_income_statement_quarter(corp, 1)

	comp_income_statement_sub_list.reverse()
	comp_income_statement_sub_list_q.reverse()
	individual_comp_income_statement_sub_list.reverse()  
	individual_comp_income_statement_sub_list_q.reverse()
	
	# Read balance sheet
	print("Read balance sheet")
	balance_sheet_sub_list		= get_balance_sheet_year(corp)
	balance_sheet_sub_list_q	= get_balance_sheet_quarter(corp)
	
	# Read individual balance sheet
	print("Read individual balance sheet")
	individual_balance_sheet_sub_list, code	= get_individual_balance_sheet_year(corp)
	individual_balance_sheet_sub_list_q		= get_individual_balance_sheet_quarter(corp)
	print(code)

	balance_sheet_sub_list.reverse()
	balance_sheet_sub_list_q.reverse()
	individual_balance_sheet_sub_list.reverse() 
	individual_balance_sheet_sub_list_q.reverse()

	# Read cashflow statement
	print("Read cashflow statement")
	cashflow_statement_sub_list		= get_cashflow_statement_year(corp)
	cashflow_statement_sub_list_q	= get_cashflow_statement_quarter(corp)

	# Read individual cashflow statement
	print("Read individual cashflow statement")
	individual_cashflow_statement_sub_list		= get_individual_cashflow_statement_year(corp)
	individual_cashflow_statement_sub_list_q	= get_individual_cashflow_statement_quarter(corp)

	cashflow_statement_sub_list.reverse()
	cashflow_statement_sub_list_q.reverse()
	individual_cashflow_statement_sub_list.reverse()
	individual_cashflow_statement_sub_list_q.reverse()

	itooza_info_list = get_info_from_itooza(code)
	
	#write_google_spreadsheet(corp, itooza_info_list, balance_sheet_sub_list, individual_balance_sheet_sub_list, income_statement_sub_list, individual_income_statement_sub_list, cashflow_statement_sub_list, individual_cashflow_statement_sub_list, balance_sheet_sub_list_q, individual_balance_sheet_sub_list_q, income_statement_sub_list_q, individual_income_statement_sub_list_q, cashflow_statement_sub_list_q, individual_cashflow_statement_sub_list_q)

	write_excel_file(corp, itooza_info_list, income_statement_sub_list_q, income_statement_sub_list, comp_income_statement_sub_list, comp_income_statement_sub_list_q, individual_income_statement_sub_list, individual_income_statement_sub_list_q, individual_comp_income_statement_sub_list, individual_comp_income_statement_sub_list_q, balance_sheet_sub_list, balance_sheet_sub_list_q, individual_balance_sheet_sub_list, individual_balance_sheet_sub_list_q, cashflow_statement_sub_list, cashflow_statement_sub_list_q, individual_cashflow_statement_sub_list, individual_cashflow_statement_sub_list_q)

def get_info_from_itooza(code):
	### Get information from itooza

	url = "http://search.itooza.com/index.htm?seName=" + code
	print(url)

	handle = None
	while handle == None:
		try:
			handle = urllib.request.urlopen(url)
			#print(handle)
		except:
			pass

	data = handle.read()
	soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

	date_list				= []
	eps_connect_list		= []
	eps_individual_list		= []
	per_list				= []
	bps_list				= []
	pbr_list				= []
	dps_list				= []
	dy_list					= []
	roe_list				= []
	net_margin_list			= []
	op_margin_list			= []
	stock_price_list		= []

	date_list_q				= []
	eps_connect_list_q		= []
	eps_individual_list_q	= []
	per_list_q				= []
	bps_list_q				= []
	pbr_list_q				= []
	dps_list_q				= []
	dy_list_q				= []
	roe_list_q				= []
	net_margin_list_q		= []
	op_margin_list_q		= []
	stock_price_list_q		= []

	# Find table
	table = soup.findAll('div', {'id':'indexTable2'})
	tr_list = table[0].findAll('tr')

	th_list = tr_list[0].findAll('th')
	for ths  in th_list:
		#print(ths.text)
		if ths.text =='N/A':
			date_list.append(0.0)
		else:
			date_list.append(ths.text)

	td_list = tr_list[1].findAll('td')
	for tds  in td_list:
		#print(tds.text)
		if tds.text =='N/A':
			eps_connect_list.append(0.0)
		else:
			eps_connect_list.append(float(tds.text.replace(',','')))

	td_list = tr_list[2].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			eps_individual_list.append(0.0)
		else:
			eps_individual_list.append(float(tds.text.replace(',','')))

	td_list = tr_list[3].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			per_list.append(0.0)
		else:
			per_list.append(float(tds.text.replace(',','')))

	td_list = tr_list[4].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			bps_list.append(0.0)
		else:
			bps_list.append(float(tds.text.replace(',','')))

	td_list = tr_list[5].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			pbr_list.append(0.0)
		else:
			pbr_list.append(float(tds.text.replace(',','')))

	td_list = tr_list[6].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			dps_list.append(0.0)
		else:
			dps_list.append(float(tds.text.replace(',','')))

	td_list = tr_list[7].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			dy_list.append(0.0)
		else:
			dy_list.append(float(tds.text.replace(',','')))

	td_list = tr_list[8].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			roe_list.append(0.0)
		else:
			roe_list.append(float(tds.text.replace(',','')))

	td_list = tr_list[9].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			net_margin_list.append(0.0)
		else:
			net_margin_list.append(float(tds.text.replace(',','')))

	td_list = tr_list[10].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			op_margin_list.append(0.0)
		else:
			op_margin_list.append(float(tds.text.replace(',','')))

	td_list = tr_list[11].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			stock_price_list.append(0.0)
		else:
			stock_price_list.append(float(tds.text.replace(',','')))

	# Find table quarter
	table = soup.findAll('div', {'id':'indexTable3'})
	tr_list = table[0].findAll('tr')

	th_list = tr_list[0].findAll('th')
	for ths  in th_list:
		if ths.text =='N/A':
			date_list_q.append(0.0)
		else:
			date_list_q.append(ths.text)

	td_list = tr_list[1].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			eps_connect_list_q.append(0.0)
		else:
			eps_connect_list_q.append(float(tds.text.replace(',','')))

	td_list = tr_list[2].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			eps_individual_list_q.append(0.0)
		else:
			eps_individual_list_q.append(float(tds.text.replace(',','')))

	td_list = tr_list[3].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			per_list_q.append(0.0)
		else:
			per_list_q.append(float(tds.text.replace(',','')))

	td_list = tr_list[4].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			bps_list_q.append(0.0)
		else:
			bps_list_q.append(float(tds.text.replace(',','')))

	td_list = tr_list[5].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			pbr_list_q.append(0.0)
		else:
			pbr_list_q.append(float(tds.text.replace(',','')))

	td_list = tr_list[6].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			dps_list_q.append(0.0)
		else:
			dps_list_q.append(float(tds.text.replace(',','')))

	td_list = tr_list[7].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			dy_list_q.append(0.0)
		else:
			dy_list_q.append(float(tds.text.replace(',','')))

	td_list = tr_list[8].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			roe_list_q.append(0.0)
		else:
			roe_list_q.append(float(tds.text.replace(',','')))

	td_list = tr_list[9].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			net_margin_list_q.append(0.0)
		else:
			net_margin_list_q.append(float(tds.text.replace(',','')))

	td_list = tr_list[10].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			op_margin_list_q.append(0.0)
		else:
			op_margin_list_q.append(float(tds.text.replace(',','')))

	td_list = tr_list[11].findAll('td')
	for tds  in td_list:
		if tds.text =='N/A':
			stock_price_list_q.append(0.0)
		else:
			stock_price_list_q.append(float(tds.text.replace(',','')))
	
	itooza_info_list = []
	
	itooza_info_list.append(date_list				)
	itooza_info_list.append(eps_connect_list		)
	itooza_info_list.append(eps_individual_list		)
	itooza_info_list.append(per_list				)
	itooza_info_list.append(bps_list				)
	itooza_info_list.append(pbr_list				)
	itooza_info_list.append(dps_list				)
	itooza_info_list.append(dy_list					)
	itooza_info_list.append(roe_list				)
	itooza_info_list.append(net_margin_list			)
	itooza_info_list.append(op_margin_list			)
	itooza_info_list.append(stock_price_list		)
	
	itooza_info_list.append(date_list_q				)
	itooza_info_list.append(eps_connect_list_q		)
	itooza_info_list.append(eps_individual_list_q	)
	itooza_info_list.append(per_list_q				)
	itooza_info_list.append(bps_list_q				)
	itooza_info_list.append(pbr_list_q				)
	itooza_info_list.append(dps_list_q				)
	itooza_info_list.append(dy_list_q				)
	itooza_info_list.append(roe_list_q				)
	itooza_info_list.append(net_margin_list_q		)
	itooza_info_list.append(op_margin_list_q		)
	itooza_info_list.append(stock_price_list_q		)

	return itooza_info_list

def write_google_spreadsheet(corp, itooza_info_list, balance_sheet_sub_list, individual_balance_sheet_sub_list, income_statement_sub_list, individual_income_statement_sub_list, cashflow_statement_sub_list, individual_cashflow_statement_sub_list, balance_sheet_sub_list_q, individual_balance_sheet_sub_list_q, income_statement_sub_list_q, individual_income_statement_sub_list_q, cashflow_statement_sub_list_q, individual_cashflow_statement_sub_list_q):
	date_list				= itooza_info_list[0]
	eps_connect_list		= itooza_info_list[1]
	eps_individual_list		= itooza_info_list[2]
	per_list				= itooza_info_list[3]
	bps_list				= itooza_info_list[4]
	pbr_list				= itooza_info_list[5]
	dps_list				= itooza_info_list[6]
	dy_list					= itooza_info_list[7]
	roe_list				= itooza_info_list[8]
	net_margin_list			= itooza_info_list[9]
	op_margin_list			= itooza_info_list[10]
	stock_price_list		= itooza_info_list[11]

	date_list_q				= itooza_info_list[12]
	eps_connect_list_q		= itooza_info_list[13]
	eps_individual_list_q	= itooza_info_list[14]
	per_list_q				= itooza_info_list[15]
	bps_list_q				= itooza_info_list[16]
	pbr_list_q				= itooza_info_list[17]
	dps_list_q				= itooza_info_list[18]
	dy_list_q				= itooza_info_list[19]
	roe_list_q				= itooza_info_list[20]
	net_margin_list_q		= itooza_info_list[21]
	op_margin_list_q		= itooza_info_list[22]
	stock_price_list_q		= itooza_info_list[23]

	# use creds to create a client to interact with the Google Drive API
	scope = ['https://spreadsheets.google.com/feeds',
			'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name('mykey.json', scope)
	client = gspread.authorize(creds)

	service = discovery.build('sheets', 'v4', credentials=creds)
	
	# The ID of the spreadsheet containing the sheet to copy.
	spreadsheet_id = '1H4x975Hj3MSRUIgCddDc42KM7-UsDdkiT62PRdCyHys' 
	
	# The ID of the sheet to copy.
	#sheet_id = 1315244837	# test worksheet
	sheet_id = 2069612066	# test.2018 worksheet
	#sheet_id = 100888209 
	
	# The ID of the spreadsheet to copy the sheet to.
	copy_sheet_to_another_spreadsheet_request_body = {'destination_spreadsheet_id': '1H4x975Hj3MSRUIgCddDc42KM7-UsDdkiT62PRdCyHys'}
	
	request = service.spreadsheets().sheets().copyTo(spreadsheetId=spreadsheet_id, sheetId=sheet_id, body=copy_sheet_to_another_spreadsheet_request_body)
	response = request.execute()

	# Find a workbook by name and open the first sheet
	# Make sure you use the right name here.
	
	#sheet = client.open("test_graph").sheet1
	spreadsheet = client.open("test_graph")
	#sheet = spreadsheet.worksheet("Copy of format")
	sheet = spreadsheet.worksheet("Copy of test.2018")
	#sheet = spreadsheet.add_worksheet(title=corp, rows="100", cols="50")
	print(type(sheet))

	sheet.update_cell(3, 2, corp)

	# 주당순이익
	cell_list = sheet.range(9,4, 9, 14)
	for k, cell in enumerate(cell_list):
		cell.value = eps_connect_list[k+1]
	sheet.update_cells(cell_list)
	
	# 주당배당금
	cell_list = sheet.range(10,4, 10, 14)
	for k, cell in enumerate(cell_list):
		cell.value = dps_list[k+1]
	sheet.update_cells(cell_list)

	# 시가배당률
	cell_list = sheet.range(11,4, 11, 14)
	for k, cell in enumerate(cell_list):
		cell.value = dy_list[k+1]
	sheet.update_cells(cell_list)

	# PER
	cell_list = sheet.range(79,4, 79, 9)
	for k, cell in enumerate(cell_list):
		cell.value = per_list[k+1]
	sheet.update_cells(cell_list)

	# PBR
	cell_list = sheet.range(80,4, 80, 9)
	for k, cell in enumerate(cell_list):
		cell.value = pbr_list[k+1]
	sheet.update_cells(cell_list)

	# 자산 총계
	cell_list = sheet.range(19, 4, 19, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['Assets'] != -1:
			cell.value = balance_sheet_sub_list[k]['Assets']
		elif individual_balance_sheet_sub_list[k]['Assets'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['Assets']
	sheet.update_cells(cell_list)

	# 현금 및 예치금
	cell_list = sheet.range(20, 4, 20, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['CashAndCashEquivalents'] != -1:
			cell.value = balance_sheet_sub_list[k]['CashAndCashEquivalents']
		elif individual_balance_sheet_sub_list[k]['CashAndCashEquivalents'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['CashAndCashEquivalents']
	sheet.update_cells(cell_list)

	# 단기금융상품
	cell_list = sheet.range(21, 4, 21, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['ShortTermDeposits'] != -1:
			cell.value = balance_sheet_sub_list[k]['ShortTermDeposits']
		elif individual_balance_sheet_sub_list[k]['ShortTermDeposits'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['ShortTermDeposits']
	sheet.update_cells(cell_list)

	# 매출채권
	cell_list = sheet.range(22, 4, 22, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['ShortTermTradeReceivable'] != -1:
			cell.value = balance_sheet_sub_list[k]['ShortTermTradeReceivable']
		elif balance_sheet_sub_list[k]['TradeAndOtherCurrentReceivables'] != -1:
			cell.value = balance_sheet_sub_list[k]['TradeAndOtherCurrentReceivables']
		elif individual_balance_sheet_sub_list[k]['ShortTermTradeReceivable'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['ShortTermTradeReceivable']
		elif individual_balance_sheet_sub_list[k]['TradeAndOtherCurrentReceivables'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['TradeAndOtherCurrentReceivables']
	sheet.update_cells(cell_list)

	# 재고자산
	cell_list = sheet.range(23, 4, 23, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['Inventories'] != -1:
			cell.value = balance_sheet_sub_list[k]['Inventories']
		elif individual_balance_sheet_sub_list[k]['Inventories'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['Inventories']
	sheet.update_cells(cell_list)

	# 유형자산
	cell_list = sheet.range(24, 4, 24, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['PropertyPlantAndEquipment'] != -1:
			cell.value = balance_sheet_sub_list[k]['PropertyPlantAndEquipment']
		elif individual_balance_sheet_sub_list[k]['PropertyPlantAndEquipment'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['PropertyPlantAndEquipment']
	sheet.update_cells(cell_list)

	# 투자부동산
	cell_list = sheet.range(25, 4, 25, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['InvestmentProperty'] != -1:
			cell.value = balance_sheet_sub_list[k]['InvestmentProperty']
		elif individual_balance_sheet_sub_list[k]['InvestmentProperty'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['InvestmentProperty']
	sheet.update_cells(cell_list)

	# 무형자산
	cell_list = sheet.range(26, 4, 26, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['IntangibleAssetsOtherThanGoodwill'] != -1:
			cell.value = balance_sheet_sub_list[k]['IntangibleAssetsOtherThanGoodwill']
		elif individual_balance_sheet_sub_list[k]['IntangibleAssetsOtherThanGoodwill'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['IntangibleAssetsOtherThanGoodwill']
	sheet.update_cells(cell_list)

	# 지분법적용 투자지분
	cell_list = sheet.range(27, 4, 27, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['InvestmentAccounted'] != -1:
			cell.value = balance_sheet_sub_list[k]['InvestmentAccounted']
		elif individual_balance_sheet_sub_list[k]['InvestmentAccounted'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['InvestmentAccounted']
	sheet.update_cells(cell_list)

	# 기타비유동자산 
	cell_list = sheet.range(28, 4, 28, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['OtherNonCurrentAssets'] != -1:
			cell.value = balance_sheet_sub_list[k]['OtherNonCurrentAssets']
		elif individual_balance_sheet_sub_list[k]['OtherNonCurrentAssets'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['OtherNonCurrentAssets']
	sheet.update_cells(cell_list)
	
	
	# 부채 총계
	cell_list = sheet.range(30, 4, 30, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['Liabilities'] != -1:
			cell.value = balance_sheet_sub_list[k]['Liabilities']
		elif individual_balance_sheet_sub_list[k]['Liabilities'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['Liabilities']
	sheet.update_cells(cell_list)
	# 단기차입금
	cell_list = sheet.range(31, 4, 31, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['ShortTermBorrowings'] != -1:
			cell.value = balance_sheet_sub_list[k]['ShortTermBorrowings']
		elif individual_balance_sheet_sub_list[k]['ShortTermBorrowings'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['ShortTermBorrowings']
	sheet.update_cells(cell_list)
	# 매입채무
	cell_list = sheet.range(34, 4, 34, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['TradeAndOtherCurrentPayables'] != -1:
			cell.value = balance_sheet_sub_list[k]['TradeAndOtherCurrentPayables']
		elif individual_balance_sheet_sub_list[k]['TradeAndOtherCurrentPayables'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['TradeAndOtherCurrentPayables']
	sheet.update_cells(cell_list)

	# 자본 총계
	cell_list = sheet.range(36, 4, 36, 8)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list[k]['Equity'] != -1:
			cell.value = balance_sheet_sub_list[k]['Equity']
		elif individual_balance_sheet_sub_list[k]['Equity'] != -1:
			cell.value = individual_balance_sheet_sub_list[k]['Equity']
	sheet.update_cells(cell_list)

	# 매출액
	cell_list = sheet.range(38, 4, 38, 8)
	for k, cell in enumerate(cell_list):
		if income_statement_sub_list[k][0] != -1:
			cell.value = income_statement_sub_list[k][0]
		elif comp_income_statement_sub_list[k][0] != -1:
			cell.value = comp_income_statement_sub_list[k][0]
		elif individual_income_statement_sub_list[k][0] != -1:
			cell.value = individual_income_statement_sub_list[k][0]
		elif individual_comp_income_statement_sub_list[k][0] != -1:
			cell.value = individual_comp_income_statement_sub_list[k][0]
	sheet.update_cells(cell_list)
	# 매출원가
	cell_list = sheet.range(39, 4, 39, 8)
	for k, cell in enumerate(cell_list):
		if income_statement_sub_list[k][1] != -1:
			cell.value = income_statement_sub_list[k][1]
		elif comp_income_statement_sub_list[k][1] != -1:
			cell.value = comp_income_statement_sub_list[k][1]
		elif individual_income_statement_sub_list[k][1] != -1:
			cell.value = individual_income_statement_sub_list[k][1]
		elif individual_comp_income_statement_sub_list[k][1] != -1:
			cell.value = individual_comp_income_statement_sub_list[k][1]
	sheet.update_cells(cell_list)
	# 판관비
	cell_list = sheet.range(40, 4, 40, 8)
	for k, cell in enumerate(cell_list):
		if income_statement_sub_list[k][3] != -1:
			cell.value = income_statement_sub_list[k][3]
		elif comp_income_statement_sub_list[k][3] != -1:
			cell.value = comp_income_statement_sub_list[k][3]
		elif individual_income_statement_sub_list[k][3] != -1:
			cell.value = individual_income_statement_sub_list[k][3]
		elif individual_comp_income_statement_sub_list[k][3] != -1:
			cell.value = individual_comp_income_statement_sub_list[k][3]
	sheet.update_cells(cell_list)
	# 영업이익
	cell_list = sheet.range(41, 4, 41, 8)
	for k, cell in enumerate(cell_list):
		if income_statement_sub_list[k][4] != -1:
			cell.value = income_statement_sub_list[k][4]
		elif comp_income_statement_sub_list[k][4] != -1:
			cell.value = comp_income_statement_sub_list[k][4]
		elif individual_income_statement_sub_list[k][4] != -1:
			cell.value = individual_income_statement_sub_list[k][4]
		elif individual_comp_income_statement_sub_list[k][4] != -1:
			cell.value = individual_comp_income_statement_sub_list[k][4]
	sheet.update_cells(cell_list)
	# 당기순이익
	cell_list = sheet.range(42, 4, 42, 8)
	for k, cell in enumerate(cell_list):
		if income_statement_sub_list[k][12] != -1:
			cell.value = income_statement_sub_list[k][12]
		elif comp_income_statement_sub_list[k][12] != -1:
			cell.value = comp_income_statement_sub_list[k][12]
		elif individual_income_statement_sub_list[k][12] != -1:
			cell.value = individual_income_statement_sub_list[k][12]
		elif individual_comp_income_statement_sub_list[k][12] != -1:
			cell.value = individual_comp_income_statement_sub_list[k][12]
	sheet.update_cells(cell_list)

	# 영업현금흐름
	cell_list = sheet.range(43, 4, 43, 8)
	for k, cell in enumerate(cell_list):
		if cashflow_statement_sub_list[k]['CashFlowsFromUsedInOperatingActivities'] != -1:
			cell.value = cashflow_statement_sub_list[k]['CashFlowsFromUsedInOperatingActivities']
		elif individual_cashflow_statement_sub_list[k]['CashFlowsFromUsedInOperatingActivities'] != -1:
			cell.value = individual_cashflow_statement_sub_list[k]['CashFlowsFromUsedInOperatingActivities']
	sheet.update_cells(cell_list)

	# 투자현금흐름
	cell_list = sheet.range(47, 4, 47, 8)
	for k, cell in enumerate(cell_list):
		if cashflow_statement_sub_list[k]['CashFlowsFromUsedInInvestingActivities'] != -1:
			cell.value = cashflow_statement_sub_list[k]['CashFlowsFromUsedInInvestingActivities']
		elif individual_cashflow_statement_sub_list[k]['CashFlowsFromUsedInInvestingActivities'] != -1:
			cell.value = individual_cashflow_statement_sub_list[k]['CashFlowsFromUsedInInvestingActivities']
	sheet.update_cells(cell_list)

	# 재무현금흐름
	cell_list = sheet.range(50, 4, 50, 8)
	for k, cell in enumerate(cell_list):
		if cashflow_statement_sub_list[k]['CashFlowsFromUsedInFinancingActivities'] != -1:
			cell.value = cashflow_statement_sub_list[k]['CashFlowsFromUsedInFinancingActivities']
		elif individual_cashflow_statement_sub_list[k]['CashFlowsFromUsedInFinancingActivities'] != -1:
			cell.value = individual_cashflow_statement_sub_list[k]['CashFlowsFromUsedInFinancingActivities']
	sheet.update_cells(cell_list)

	### 분기별 실적 파트

	# 주당순이익(연결)
	cell_list = sheet.range(84,4, 84, 11)
	for k, cell in enumerate(cell_list):
		cell.value = pbr_list_q[k+1]
	sheet.update_cells(cell_list)
	
	# 주당순자산
	cell_list = sheet.range(85,4, 85, 11)
	for k, cell in enumerate(cell_list):
		cell.value = pbr_list_q[k+1]
	sheet.update_cells(cell_list)

	# PBR
	cell_list = sheet.range(87,4, 87, 11)
	for k, cell in enumerate(cell_list):
		cell.value = pbr_list_q[k+1]
	sheet.update_cells(cell_list)

	# 자산 총계
	cell_list = sheet.range(92, 4, 92, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['Assets'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['Assets']
		elif individual_balance_sheet_sub_list_q[k]['Assets'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['Assets']
	sheet.update_cells(cell_list)

	# 현금 및 예치금
	cell_list = sheet.range(93, 4, 93, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['CashAndCashEquivalents'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['CashAndCashEquivalents']
		elif individual_balance_sheet_sub_list_q[k]['CashAndCashEquivalents'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['CashAndCashEquivalents']
	sheet.update_cells(cell_list)

	# 단기금융상품
	cell_list = sheet.range(94, 4, 94, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['ShortTermDeposits'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['ShortTermDeposits']
		elif individual_balance_sheet_sub_list_q[k]['ShortTermDeposits'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['ShortTermDeposits']
	sheet.update_cells(cell_list)

	# 매출채권
	cell_list = sheet.range(95, 4, 95, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['ShortTermTradeReceivable'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['ShortTermTradeReceivable']
		elif balance_sheet_sub_list_q[k]['TradeAndOtherCurrentReceivables'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['TradeAndOtherCurrentReceivables']
		elif individual_balance_sheet_sub_list_q[k]['ShortTermTradeReceivable'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['ShortTermTradeReceivable']
		elif individual_balance_sheet_sub_list_q[k]['TradeAndOtherCurrentReceivables'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['TradeAndOtherCurrentReceivables']
	sheet.update_cells(cell_list)

	# 재고자산
	cell_list = sheet.range(96, 4, 96, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['Inventories'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['Inventories']
		elif individual_balance_sheet_sub_list_q[k]['Inventories'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['Inventories']
	sheet.update_cells(cell_list)

	# 유형자산
	cell_list = sheet.range(97, 4, 97, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['PropertyPlantAndEquipment'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['PropertyPlantAndEquipment']
		elif individual_balance_sheet_sub_list_q[k]['PropertyPlantAndEquipment'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['PropertyPlantAndEquipment']
	sheet.update_cells(cell_list)

	# 투자부동산
	cell_list = sheet.range(98, 4, 98, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['InvestmentProperty'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['InvestmentProperty']
		elif individual_balance_sheet_sub_list_q[k]['InvestmentProperty'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['InvestmentProperty']
	sheet.update_cells(cell_list)

	# 무형자산
	cell_list = sheet.range(99, 4, 99, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['IntangibleAssetsOtherThanGoodwill'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['IntangibleAssetsOtherThanGoodwill']
		elif individual_balance_sheet_sub_list_q[k]['IntangibleAssetsOtherThanGoodwill'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['IntangibleAssetsOtherThanGoodwill']
	sheet.update_cells(cell_list)

	# 지분법적용 투자지분
	cell_list = sheet.range(100, 4, 100, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['InvestmentAccounted'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['InvestmentAccounted']
		elif individual_balance_sheet_sub_list_q[k]['InvestmentAccounted'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['InvestmentAccounted']
	sheet.update_cells(cell_list)

	# 기타비유동자산 
	cell_list = sheet.range(101, 4, 101, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['OtherNonCurrentAssets'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['OtherNonCurrentAssets']
		elif individual_balance_sheet_sub_list_q[k]['OtherNonCurrentAssets'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['OtherNonCurrentAssets']
	sheet.update_cells(cell_list)
	
	# 부채 총계
	cell_list = sheet.range(103, 4, 103, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['Liabilities'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['Liabilities']
		elif individual_balance_sheet_sub_list_q[k]['Liabilities'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['Liabilities']
	sheet.update_cells(cell_list)
	# 단기차입금
	cell_list = sheet.range(104, 4, 104, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['ShortTermBorrowings'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['ShortTermBorrowings']
		elif individual_balance_sheet_sub_list_q[k]['ShortTermBorrowings'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['ShortTermBorrowings']
	sheet.update_cells(cell_list)
	# 매입채무
	cell_list = sheet.range(107, 4, 107, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['TradeAndOtherCurrentPayables'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['TradeAndOtherCurrentPayables']
		elif individual_balance_sheet_sub_list_q[k]['TradeAndOtherCurrentPayables'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['TradeAndOtherCurrentPayables']
	sheet.update_cells(cell_list)

	# 자본 총계
	cell_list = sheet.range(109, 4, 109, 11)
	for k, cell in enumerate(cell_list):
		if balance_sheet_sub_list_q[k]['Equity'] != -1:
			cell.value = balance_sheet_sub_list_q[k]['Equity']
		elif individual_balance_sheet_sub_list_q[k]['Equity'] != -1:
			cell.value = individual_balance_sheet_sub_list_q[k]['Equity']
	sheet.update_cells(cell_list)

	# 매출액
	cell_list = sheet.range(111, 4, 111, 11)
	for k, cell in enumerate(cell_list):
		if income_statement_sub_list_q[k][0] != -1:
			cell.value = income_statement_sub_list_q[k][0]
		elif comp_income_statement_sub_list_q[k][0] != -1:
			cell.value = comp_income_statement_sub_list_q[k][0]
		elif individual_income_statement_sub_list_q[k][0] != -1:
			cell.value = individual_income_statement_sub_list_q[k][0]
		elif individual_comp_income_statement_sub_list_q[k][0] != -1:
			cell.value = individual_comp_income_statement_sub_list_q[k][0]
	sheet.update_cells(cell_list)
	# 매출원가
	cell_list = sheet.range(112, 4, 112, 11)
	for k, cell in enumerate(cell_list):
		if income_statement_sub_list_q[k][1] != -1:
			cell.value = income_statement_sub_list_q[k][1]
		elif comp_income_statement_sub_list_q[k][1] != -1:
			cell.value = comp_income_statement_sub_list_q[k][1]
		elif individual_income_statement_sub_list_q[k][1] != -1:
			cell.value = individual_income_statement_sub_list_q[k][1]
		elif individual_comp_income_statement_sub_list_q[k][1] != -1:
			cell.value = individual_comp_income_statement_sub_list_q[k][1]
	sheet.update_cells(cell_list)
	# 매출총이익
	cell_list = sheet.range(113, 4, 113, 11)
	for k, cell in enumerate(cell_list):
		if income_statement_sub_list_q[k][2] != -1:
			cell.value = income_statement_sub_list_q[k][1]
		elif comp_income_statement_sub_list_q[k][2] != -1:
			cell.value = comp_income_statement_sub_list_q[k][1]
		elif individual_income_statement_sub_list_q[k][2] != -1:
			cell.value = individual_income_statement_sub_list_q[k][1]
		elif individual_comp_income_statement_sub_list_q[k][2] != -1:
			cell.value = individual_comp_income_statement_sub_list_q[k][1]
	sheet.update_cells(cell_list)
	# 판관비
	cell_list = sheet.range(114, 4, 114, 11)
	for k, cell in enumerate(cell_list):
		if income_statement_sub_list_q[k][3] != -1:
			cell.value = income_statement_sub_list_q[k][3]
		elif comp_income_statement_sub_list_q[k][3] != -1:
			cell.value = comp_income_statement_sub_list_q[k][3]
		elif individual_income_statement_sub_list_q[k][3] != -1:
			cell.value = individual_income_statement_sub_list_q[k][3]
		elif individual_comp_income_statement_sub_list_q[k][3] != -1:
			cell.value = individual_comp_income_statement_sub_list_q[k][3]
	sheet.update_cells(cell_list)
	# 영업이익
	cell_list = sheet.range(115, 4, 115, 11)
	for k, cell in enumerate(cell_list):
		if income_statement_sub_list_q[k][4] != -1:
			cell.value = income_statement_sub_list_q[k][4]
		elif comp_income_statement_sub_list_q[k][4] != -1:
			cell.value = comp_income_statement_sub_list_q[k][4]
		elif individual_income_statement_sub_list_q[k][4] != -1:
			cell.value = individual_income_statement_sub_list_q[k][4]
		elif individual_comp_income_statement_sub_list_q[k][4] != -1:
			cell.value = individual_comp_income_statement_sub_list_q[k][4]
	sheet.update_cells(cell_list)
	# 당기순이익
	cell_list = sheet.range(116, 4, 116, 11)
	for k, cell in enumerate(cell_list):
		if income_statement_sub_list_q[k][12] != -1:
			cell.value = income_statement_sub_list_q[k][12]
		elif comp_income_statement_sub_list_q[k][12] != -1:
			cell.value = comp_income_statement_sub_list_q[k][12]
		elif individual_income_statement_sub_list_q[k][12] != -1:
			cell.value = individual_income_statement_sub_list_q[k][12]
		elif individual_comp_income_statement_sub_list_q[k][12] != -1:
			cell.value = individual_comp_income_statement_sub_list_q[k][12]
	sheet.update_cells(cell_list)

	# 영업현금흐름
	cell_list = sheet.range(117, 4, 117, 11)
	for k, cell in enumerate(cell_list):
		if cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInOperatingActivities'] != -1:
			cell.value = cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInOperatingActivities']
		elif individual_cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInOperatingActivities'] != -1:
			cell.value = individual_cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInOperatingActivities']
	sheet.update_cells(cell_list)

	# 투자현금흐름
	cell_list = sheet.range(121, 4, 121, 11)
	for k, cell in enumerate(cell_list):
		if cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInInvestingActivities'] != -1:
			cell.value = cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInInvestingActivities']
		elif individual_cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInInvestingActivities'] != -1:
			cell.value = individual_cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInInvestingActivities']
	sheet.update_cells(cell_list)

	# 재무현금흐름
	cell_list = sheet.range(124, 4, 124, 11)
	for k, cell in enumerate(cell_list):
		if cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInFinancingActivities'] != -1:
			cell.value = cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInFinancingActivities']
		elif individual_cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInFinancingActivities'] != -1:
			cell.value = individual_cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInFinancingActivities']
	sheet.update_cells(cell_list)

	sheet.update_title(corp)

def	write_excel_file(corp, itooza_info_list, income_statement_sub_list_q, income_statement_sub_list, comp_income_statement_sub_list, comp_income_statement_sub_list_q, individual_income_statement_sub_list, individual_income_statement_sub_list_q, individual_comp_income_statement_sub_list, individual_comp_income_statement_sub_list_q, balance_sheet_sub_list, balance_sheet_sub_list_q, individual_balance_sheet_sub_list, individual_balance_sheet_sub_list_q, cashflow_statement_sub_list, cashflow_statement_sub_list_q, individual_cashflow_statement_sub_list, individual_cashflow_statement_sub_list_q):

	date_list				= itooza_info_list[0]
	eps_connect_list		= itooza_info_list[1]
	eps_individual_list		= itooza_info_list[2]
	per_list				= itooza_info_list[3]
	bps_list				= itooza_info_list[4]
	pbr_list				= itooza_info_list[5]
	dps_list				= itooza_info_list[6]
	dy_list					= itooza_info_list[7]
	roe_list				= itooza_info_list[8]
	net_margin_list			= itooza_info_list[9]
	op_margin_list			= itooza_info_list[10]
	stock_price_list		= itooza_info_list[11]

	date_list_q				= itooza_info_list[12]
	eps_connect_list_q		= itooza_info_list[13]
	eps_individual_list_q	= itooza_info_list[14]
	per_list_q				= itooza_info_list[15]
	bps_list_q				= itooza_info_list[16]
	pbr_list_q				= itooza_info_list[17]
	dps_list_q				= itooza_info_list[18]
	dy_list_q				= itooza_info_list[19]
	roe_list_q				= itooza_info_list[20]
	net_margin_list_q		= itooza_info_list[21]
	op_margin_list_q		= itooza_info_list[22]
	stock_price_list_q		= itooza_info_list[23]

	# Write excel file
	workbook_name = "result_"+corp+".xlsx"
	cur_dir = os.getcwd()
	if os.path.isfile(os.path.join(cur_dir, workbook_name)):
		os.remove(os.path.join(cur_dir, workbook_name))
	workbook = xlsxwriter.Workbook(workbook_name)

	filter_format = workbook.add_format({'bold':True,
										'fg_color': '#D7E4BC'
										})
	filter_format2 = workbook.add_format({'bold':True
										})

	percent_format = workbook.add_format({'num_format': '0.00%'})

	roe_format = workbook.add_format({'bold':True,
									  'underline': True,
									  'num_format': '0.00%'})

	num_format = workbook.add_format({'num_format':'0.00'})
	num2_format = workbook.add_format({'num_format':'#,##0'})
	num3_format = workbook.add_format({'num_format':'#,##0.00',
									  'fg_color':'#FCE4D6'})

	worksheet_1 = workbook.add_worksheet('연결_재무제표_연도')
	
	worksheet_1.write(0, 0, "연결 손익계산서", filter_format2)
	worksheet_1.set_column('A:A', 30)
	worksheet_1.write(1, 0, "매출액")
	worksheet_1.write(2, 0, "매출원가")
	worksheet_1.write(3, 0, "매출총이익")
	worksheet_1.write(4, 0, "판매비와 관리비")
	worksheet_1.write(5, 0, "영업이익")
	worksheet_1.write(6, 0, "기타수익")
	worksheet_1.write(7, 0, "기타비용")
	worksheet_1.write(8, 0, "금융수익")
	worksheet_1.write(9, 0, "금융비용")
	worksheet_1.write(10, 0, "지분법이익")
	worksheet_1.write(11, 0, "법인세비용차감전순이익")
	worksheet_1.write(12, 0, "법인세비용")
	worksheet_1.write(13, 0, "당기순이익")
	worksheet_1.write(14, 0, "기본주당순이익")
	
	worksheet_1.set_column('B:B', 15)
	worksheet_1.write(0, 1,		"2017년", filter_format)
	worksheet_1.write(0, 2, 	"2016년", filter_format)
	worksheet_1.write(0, 3, 	"2015년", filter_format)
	worksheet_1.write(0, 4, 	"2014년", filter_format)
	worksheet_1.write(0, 5,		"2013년", filter_format)
	
	for k in range(len(income_statement_sub_list)):
		for l in range(len(income_statement_sub_list[k])):
			worksheet_1.write(l+1, k+1, income_statement_sub_list[k][l], num2_format)
	
	worksheet_2 = workbook.add_worksheet('연결_재무제표_분기')
	worksheet_2.write(0, 0, "연결 손익계산서", filter_format2)
	worksheet_2.set_column('A:A', 30)
	worksheet_2.write(1, 0, "매출액")
	worksheet_2.write(2, 0, "매출원가")
	worksheet_2.write(3, 0, "매출총이익")
	worksheet_2.write(4, 0, "판매비와 관리비")
	worksheet_2.write(5, 0, "영업이익")
	worksheet_2.write(6, 0, "기타수익")
	worksheet_2.write(7, 0, "기타비용")
	worksheet_2.write(8, 0, "금융수익")
	worksheet_2.write(9, 0, "금융비용")
	worksheet_2.write(10, 0, "지분법이익")
	worksheet_2.write(11, 0, "법인세비용차감전순이익")
	worksheet_2.write(12, 0, "법인세비용")
	worksheet_2.write(13, 0, "당기순이익")
	worksheet_2.write(14, 0, "기본주당순이익")
	
	worksheet_2.set_column('B:B', 15)
	worksheet_2.write(0, 1,		"2017.4Q", filter_format)
	worksheet_2.write(0, 2,		"2017.3Q", filter_format)
	worksheet_2.write(0, 3, 	"2017.2Q", filter_format)
	worksheet_2.write(0, 4, 	"2017.1Q", filter_format)
	worksheet_2.write(0, 5, 	"2016.4Q", filter_format)
	worksheet_2.write(0, 6, 	"2016.3Q", filter_format)
	worksheet_2.write(0, 7, 	"2016.2Q", filter_format)
	worksheet_2.write(0, 8, 	"2016.1Q", filter_format)
	
	for k in range(len(income_statement_sub_list_q)):
		for l in range(len(income_statement_sub_list_q[k])):
			worksheet_2.write(l+1, k+1, income_statement_sub_list_q[k][l], num2_format)
	
	#worksheet_11 = workbook.add_worksheet('연결_포괄손익계산서_year')
	offset = len(income_statement_sub_list[0]) + 2

	worksheet_1.write(offset + 0, 0, "연결 포괄손익계산서", filter_format2)
	worksheet_1.set_column('A:A', 30)
	worksheet_1.write(offset + 1, 0, "매출액")
	worksheet_1.write(offset + 2, 0, "매출원가")
	worksheet_1.write(offset + 3, 0, "매출총이익")
	worksheet_1.write(offset + 4, 0, "판매비와 관리비")
	worksheet_1.write(offset + 5, 0, "영업이익")
	worksheet_1.write(offset + 6, 0, "기타수익")
	worksheet_1.write(offset + 7, 0, "기타비용")
	worksheet_1.write(offset + 8, 0, "금융수익")
	worksheet_1.write(offset + 9, 0, "금융비용")
	worksheet_1.write(offset + 10, 0, "지분법이익")
	worksheet_1.write(offset + 11, 0, "법인세비용차감전순이익")
	worksheet_1.write(offset + 12, 0, "법인세비용")
	worksheet_1.write(offset + 13, 0, "당기순이익")
	worksheet_1.write(offset + 14, 0, "기본주당순이익")
	
	worksheet_1.write(offset + 0, 1,	"2017년", filter_format)
	worksheet_1.write(offset + 0, 2, 	"2016년", filter_format)
	worksheet_1.write(offset + 0, 3, 	"2015년", filter_format)
	worksheet_1.write(offset + 0, 4, 	"2014년", filter_format)
	worksheet_1.write(offset + 0, 5,	"2013년", filter_format)
	
	for k in range(len(comp_income_statement_sub_list)):
		for l in range(len(comp_income_statement_sub_list[k])):
			worksheet_1.write(offset+l+1, k+1, comp_income_statement_sub_list[k][l], num2_format)
	
	#worksheet_12 = workbook.add_worksheet('연결_포괄손익계산서_분기')
	worksheet_2.write(offset + 0, 0, "연결 포괄손익계산서", filter_format2)
	worksheet_2.set_column('A:A', 30)
	worksheet_2.write(offset + 1, 0, "매출액")
	worksheet_2.write(offset + 2, 0, "매출원가")
	worksheet_2.write(offset + 3, 0, "매출총이익")
	worksheet_2.write(offset + 4, 0, "판매비와 관리비")
	worksheet_2.write(offset + 5, 0, "영업이익")
	worksheet_2.write(offset + 6, 0, "기타수익")
	worksheet_2.write(offset + 7, 0, "기타비용")
	worksheet_2.write(offset + 8, 0, "금융수익")
	worksheet_2.write(offset + 9, 0, "금융비용")
	worksheet_2.write(offset + 10, 0, "지분법이익")
	worksheet_2.write(offset + 11, 0, "법인세비용차감전순이익")
	worksheet_2.write(offset + 12, 0, "법인세비용")
	worksheet_2.write(offset + 13, 0, "당기순이익")
	worksheet_2.write(offset + 14, 0, "기본주당순이익")
	
	worksheet_2.write(offset + 0, 1,	"2017.4Q", filter_format)
	worksheet_2.write(offset + 0, 2,	"2017.3Q", filter_format)
	worksheet_2.write(offset + 0, 3, 	"2017.2Q", filter_format)
	worksheet_2.write(offset + 0, 4, 	"2017.1Q", filter_format)
	worksheet_2.write(offset + 0, 5, 	"2016.4Q", filter_format)
	worksheet_2.write(offset + 0, 6, 	"2016.3Q", filter_format)
	worksheet_2.write(offset + 0, 7, 	"2016.2Q", filter_format)
	worksheet_2.write(offset + 0, 8, 	"2016.1Q", filter_format)
	
	for k in range(len(comp_income_statement_sub_list_q)):
		for l in range(len(comp_income_statement_sub_list_q[k])):
			worksheet_2.write(offset+l+1, k+1, comp_income_statement_sub_list_q[k][l], num2_format)
	
	worksheet_3 = workbook.add_worksheet('개별_재무제표_연도')
	
	worksheet_3.write(0, 0, "개별 손익계산서", filter_format2)
	worksheet_3.set_column('A:A', 30)
	worksheet_3.write(1, 0, "매출액")
	worksheet_3.write(2, 0, "매출원가")
	worksheet_3.write(3, 0, "매출총이익")
	worksheet_3.write(4, 0, "판매비와 관리비")
	worksheet_3.write(5, 0, "영업이익")
	worksheet_3.write(6, 0, "기타수익")
	worksheet_3.write(7, 0, "기타비용")
	worksheet_3.write(8, 0, "금융수익")
	worksheet_3.write(9, 0, "금융비용")
	worksheet_3.write(10, 0, "지분법이익")
	worksheet_3.write(11, 0, "법인세비용차감전순이익")
	worksheet_3.write(12, 0, "법인세비용")
	worksheet_3.write(13, 0, "당기순이익")
	worksheet_3.write(14, 0, "기본주당순이익")
	
	worksheet_3.set_column('B:B', 15)
	worksheet_3.write(0, 1,		"2017년", filter_format)
	worksheet_3.write(0, 2, 	"2016년", filter_format)
	worksheet_3.write(0, 3, 	"2015년", filter_format)
	worksheet_3.write(0, 4, 	"2014년", filter_format)
	worksheet_3.write(0, 5,		"2013년", filter_format)
	
	for k in range(len(individual_income_statement_sub_list)):
		for l in range(len(individual_income_statement_sub_list[k])):
			worksheet_3.write(l+1, k+1, individual_income_statement_sub_list[k][l], num2_format)
	
	worksheet_4 = workbook.add_worksheet('개별_재무제표_분기')
	worksheet_4.write(0, 0, "개별 손익계산서", filter_format2)
	worksheet_4.set_column('A:A', 30)
	worksheet_4.write(1, 0, "매출액")
	worksheet_4.write(2, 0, "매출원가")
	worksheet_4.write(3, 0, "매출총이익")
	worksheet_4.write(4, 0, "판매비와 관리비")
	worksheet_4.write(5, 0, "영업이익")
	worksheet_4.write(6, 0, "기타수익")
	worksheet_4.write(7, 0, "기타비용")
	worksheet_4.write(8, 0, "금융수익")
	worksheet_4.write(9, 0, "금융비용")
	worksheet_4.write(10, 0, "지분법이익")
	worksheet_4.write(11, 0, "법인세비용차감전순이익")
	worksheet_4.write(12, 0, "법인세비용")
	worksheet_4.write(13, 0, "당기순이익")
	worksheet_4.write(14, 0, "기본주당순이익")
	
	worksheet_4.set_column('B:B', 15)
	worksheet_4.write(0, 1,		"2017.4Q", filter_format)
	worksheet_4.write(0, 2,		"2017.3Q", filter_format)
	worksheet_4.write(0, 3, 	"2017.2Q", filter_format)
	worksheet_4.write(0, 4, 	"2017.1Q", filter_format)
	worksheet_4.write(0, 5, 	"2016.4Q", filter_format)
	worksheet_4.write(0, 6, 	"2016.3Q", filter_format)
	worksheet_4.write(0, 7, 	"2016.2Q", filter_format)
	worksheet_4.write(0, 8, 	"2016.1Q", filter_format)
	
	for k in range(len(individual_income_statement_sub_list_q)):
		for l in range(len(individual_income_statement_sub_list_q[k])):
			worksheet_4.write(l+1, k+1, individual_income_statement_sub_list_q[k][l], num2_format)
	
	#worksheet_13 = workbook.add_worksheet('개별_포괄손익계산서_year')
	
	worksheet_3.write(offset + 0, 0, "개별 포괄손익계산서", filter_format2)
	worksheet_3.set_column('A:A', 30)
	worksheet_3.write(offset + 1, 0, "매출액")
	worksheet_3.write(offset + 2, 0, "매출원가")
	worksheet_3.write(offset + 3, 0, "매출총이익")
	worksheet_3.write(offset + 4, 0, "판매비와 관리비")
	worksheet_3.write(offset + 5, 0, "영업이익")
	worksheet_3.write(offset + 6, 0, "기타수익")
	worksheet_3.write(offset + 7, 0, "기타비용")
	worksheet_3.write(offset + 8, 0, "금융수익")
	worksheet_3.write(offset + 9, 0, "금융비용")
	worksheet_3.write(offset + 10, 0, "지분법이익")
	worksheet_3.write(offset + 11, 0, "법인세비용차감전순이익")
	worksheet_3.write(offset + 12, 0, "법인세비용")
	worksheet_3.write(offset + 13, 0, "당기순이익")
	worksheet_3.write(offset + 14, 0, "기본주당순이익")
	
	worksheet_3.write(offset + 0, 1,	"2017년", filter_format)
	worksheet_3.write(offset + 0, 2, 	"2016년", filter_format)
	worksheet_3.write(offset + 0, 3, 	"2015년", filter_format)
	worksheet_3.write(offset + 0, 4, 	"2014년", filter_format)
	worksheet_3.write(offset + 0, 5,	"2013년", filter_format)
	
	for k in range(len(individual_comp_income_statement_sub_list)):
		for l in range(len(individual_comp_income_statement_sub_list[k])):
			worksheet_3.write(offset+l+1, k+1, individual_comp_income_statement_sub_list[k][l], num2_format)
	
	#worksheet_14 = workbook.add_worksheet('개별_포괄손익계산서_분기')
	
	worksheet_4.write(offset + 0, 0, "개별 포괄손익계산서", filter_format2)
	worksheet_4.set_column('A:A', 30)
	worksheet_4.write(offset + 1, 0, "매출액")
	worksheet_4.write(offset + 2, 0, "매출원가")
	worksheet_4.write(offset + 3, 0, "매출총이익")
	worksheet_4.write(offset + 4, 0, "판매비와 관리비")
	worksheet_4.write(offset + 5, 0, "영업이익")
	worksheet_4.write(offset + 6, 0, "기타수익")
	worksheet_4.write(offset + 7, 0, "기타비용")
	worksheet_4.write(offset + 8, 0, "금융수익")
	worksheet_4.write(offset + 9, 0, "금융비용")
	worksheet_4.write(offset + 10, 0, "지분법이익")
	worksheet_4.write(offset + 11, 0, "법인세비용차감전순이익")
	worksheet_4.write(offset + 12, 0, "법인세비용")
	worksheet_4.write(offset + 13, 0, "당기순이익")
	worksheet_4.write(offset + 14, 0, "기본주당순이익")
	
	worksheet_4.write(offset + 0, 1,	"2017.4Q", filter_format)
	worksheet_4.write(offset + 0, 2,	"2017.3Q", filter_format)
	worksheet_4.write(offset + 0, 3, 	"2017.2Q", filter_format)
	worksheet_4.write(offset + 0, 4, 	"2017.1Q", filter_format)
	worksheet_4.write(offset + 0, 5, 	"2016.4Q", filter_format)
	worksheet_4.write(offset + 0, 6, 	"2016.3Q", filter_format)
	worksheet_4.write(offset + 0, 7, 	"2016.2Q", filter_format)
	worksheet_4.write(offset + 0, 8, 	"2016.1Q", filter_format)
	
	for k in range(len(individual_comp_income_statement_sub_list_q)):
		for l in range(len(individual_comp_income_statement_sub_list_q[k])):
			worksheet_4.write(offset+l+1, k+1, individual_comp_income_statement_sub_list_q[k][l], num2_format)
	
	#workbook_name = "test.xlsx"
	#cur_dir = os.getcwd()
	#if os.path.isfile(os.path.join(cur_dir, workbook_name)):
	#	os.remove(os.path.join(cur_dir, workbook_name))
	#workbook = xlsxwriter.Workbook(workbook_name)
	#worksheet_1 = workbook.add_worksheet('RAW')
	#for k in range(len(raw_data_list[3])):
	#	for l in range(len(raw_data_list[3][k])):
	#		worksheet_1.write(k, l, raw_data_list[3][k][l])

	#worksheet_5 = workbook.add_worksheet('연결_재무상태표_year')

	offset = offset + len(comp_income_statement_sub_list[0]) + 2

	worksheet_1.write(offset + 0, 0, "연결 재무상태표", filter_format2)
	worksheet_1.set_column('A:A', 30)
	worksheet_1.write(offset + 1, 0, "유동자산")
	worksheet_1.write(offset + 2, 0, "  현금및현금성자산")
	worksheet_1.write(offset + 3, 0, "  단기금융상품")
	worksheet_1.write(offset + 4, 0, "  기타유동금융자산")
	worksheet_1.write(offset + 5, 0, "  매출채권")
	worksheet_1.write(offset + 6, 0, "  기타유동채권")
	worksheet_1.write(offset + 7, 0, "  단기미수금")
	worksheet_1.write(offset + 8, 0, "  단기선급금")
	worksheet_1.write(offset + 9, 0, "  단기선급비용")
	worksheet_1.write(offset + 10, 0, "  재고자산")
	worksheet_1.write(offset + 11, 0, "  기타유동자산")
	worksheet_1.write(offset + 12, 0, "  당기법인세자산")
	worksheet_1.write(offset + 13, 0, "  매각예정분류자산")

	worksheet_1.write(offset + 14, 0, "비유동자산")
	worksheet_1.write(offset + 15, 0, "  장기금융상품")
	worksheet_1.write(offset + 16, 0, "  기타비유동금융자산")
	worksheet_1.write(offset + 17, 0, "  장기매출채권 및 기타비유동채권")
	worksheet_1.write(offset + 18, 0, "  장기매출채권")
	worksheet_1.write(offset + 19, 0, "  유형자산")
	worksheet_1.write(offset + 20, 0, "  투자부동산")
	worksheet_1.write(offset + 21, 0, "  영업권")
	worksheet_1.write(offset + 22, 0, "  영업권 이외의 무형자산")
	worksheet_1.write(offset + 23, 0, "  지분법적용 투자지분")
	worksheet_1.write(offset + 24, 0, "  이연법인세자산")
	worksheet_1.write(offset + 25, 0, "  기타비유동자산")
	worksheet_1.write(offset + 26, 0, "자산총계", filter_format2)

	worksheet_1.write(offset + 27, 0, "유동부채")
	worksheet_1.write(offset + 28, 0, "  매입채무 및 기타유동채무")
	worksheet_1.write(offset + 29, 0, "  단기매입채무")
	worksheet_1.write(offset + 30, 0, "  단기미지급금")
	worksheet_1.write(offset + 31, 0, "  단기선수금")
	worksheet_1.write(offset + 32, 0, "  단기예수금")
	worksheet_1.write(offset + 33, 0, "  단기차입금")
	worksheet_1.write(offset + 34, 0, "  유동성장기차입금")
	worksheet_1.write(offset + 35, 0, "  당기법인세부채")
	worksheet_1.write(offset + 36, 0, "  기타유동금융부채")
	worksheet_1.write(offset + 37, 0, "  유동충당부채")
	worksheet_1.write(offset + 38, 0, "  기타유동부채")
	worksheet_1.write(offset + 39, 0, "  매각예정분류부채")
	
	worksheet_1.write(offset + 40, 0, "비유동부채")
	worksheet_1.write(offset + 41, 0, "  장기매입채무 및 기타비유동채무")
	worksheet_1.write(offset + 42, 0, "  사채")
	worksheet_1.write(offset + 43, 0, "  장기차입금")
	worksheet_1.write(offset + 44, 0, "  기타비유동금융부채")
	worksheet_1.write(offset + 45, 0, "  비유동충당부채")
	worksheet_1.write(offset + 46, 0, "  퇴직급여부채")
	worksheet_1.write(offset + 47, 0, "  이연법인세부채")
	worksheet_1.write(offset + 48, 0, "  기타비유동부채")
	worksheet_1.write(offset + 49, 0, "부채총계", filter_format2)
	
	worksheet_1.write(offset + 50, 0, "  자본금")
	worksheet_1.write(offset + 51, 0, "  자본잉여금")
	worksheet_1.write(offset + 52, 0, "  이익잉여금")
	worksheet_1.write(offset + 53, 0, "자본총계", filter_format2)
	
	worksheet_1.write(offset + 0, 1,	"2017년", filter_format)
	worksheet_1.write(offset + 0, 2, 	"2016년", filter_format)
	worksheet_1.write(offset + 0, 3, 	"2015년", filter_format)
	worksheet_1.write(offset + 0, 4, 	"2014년", filter_format)
	worksheet_1.write(offset + 0, 5,	"2013년", filter_format)
	
	for k in range(len(balance_sheet_sub_list)):
	
		worksheet_1.write(offset + 1, k+1, balance_sheet_sub_list[k]['CurrentAssets'], num2_format)						
		worksheet_1.write(offset + 2, k+1, balance_sheet_sub_list[k]['CashAndCashEquivalents'], num2_format)			
		worksheet_1.write(offset + 3, k+1, balance_sheet_sub_list[k]['ShortTermDeposits'], num2_format)					
		worksheet_1.write(offset + 4, k+1, balance_sheet_sub_list[k]['OtherCurrentFinancialAssets'], num2_format)	
		worksheet_1.write(offset + 5, k+1, balance_sheet_sub_list[k]['ShortTermTradeReceivable'], num2_format)	
		worksheet_1.write(offset + 6, k+1, balance_sheet_sub_list[k]['TradeAndOtherCurrentReceivables'], num2_format)	
		worksheet_1.write(offset + 7, k+1, balance_sheet_sub_list[k]['ShortTermOtherReceivables'], num2_format)	
		worksheet_1.write(offset + 8, k+1, balance_sheet_sub_list[k]['ShortTermAdvancePayments'], num2_format)	
		worksheet_1.write(offset + 9, k+1, balance_sheet_sub_list[k]['ShortTermPrepaidExpenses'], num2_format)	
		worksheet_1.write(offset + 10, k+1, balance_sheet_sub_list[k]['Inventories'], num2_format)						
		worksheet_1.write(offset + 11, k+1, balance_sheet_sub_list[k]['OtherCurrentNonfinancialAssets'], num2_format)	
		worksheet_1.write(offset + 12, k+1, balance_sheet_sub_list[k]['CurrentTaxAssets'], num2_format)	
		worksheet_1.write(offset + 13, k+1, balance_sheet_sub_list[k]['NoncurrentAssetsOrDisposal'], num2_format)	
		
		worksheet_1.write(offset + 14, k+1, balance_sheet_sub_list[k]['NoncurrentAssets'], num2_format)					
		worksheet_1.write(offset + 15, k+1, balance_sheet_sub_list[k]['LongTermDeposits'], num2_format)						
		worksheet_1.write(offset + 16, k+1, balance_sheet_sub_list[k]['OtherNoncurrentFinancialAssets'], num2_format)						
		worksheet_1.write(offset + 17, k+1, balance_sheet_sub_list[k]['LongTermTradeAndOther'], num2_format)						
		worksheet_1.write(offset + 18, k+1, balance_sheet_sub_list[k]['LongTermTradeReceivablesGross'], num2_format)						
		worksheet_1.write(offset + 19, k+1, balance_sheet_sub_list[k]['PropertyPlantAndEquipment'], num2_format)			
		worksheet_1.write(offset + 20, k+1, balance_sheet_sub_list[k]['InvestmentProperty'], num2_format)				
		worksheet_1.write(offset + 21, k+1, balance_sheet_sub_list[k]['GoodwillGross'], num2_format)						
		worksheet_1.write(offset + 22, k+1, balance_sheet_sub_list[k]['IntangibleAssetsOtherThanGoodwill'], num2_format)						
		worksheet_1.write(offset + 23, k+1, balance_sheet_sub_list[k]['InvestmentAccounted'], num2_format)						
		worksheet_1.write(offset + 24, k+1, balance_sheet_sub_list[k]['DeferredTaxAssets'], num2_format)					
		worksheet_1.write(offset + 25, k+1, balance_sheet_sub_list[k]['OtherNonCurrentAssets'], num2_format)						
		worksheet_1.write(offset + 26, k+1, balance_sheet_sub_list[k]['Assets'], num2_format)							
		
		worksheet_1.write(offset + 27, k+1, balance_sheet_sub_list[k]['CurrentLiabilities'], num2_format)				
		worksheet_1.write(offset + 28, k+1, balance_sheet_sub_list[k]['TradeAndOtherCurrentPayables'], num2_format)		
		worksheet_1.write(offset + 29, k+1, balance_sheet_sub_list[k]['ShortTermTradePayables'], num2_format)						
		worksheet_1.write(offset + 30, k+1, balance_sheet_sub_list[k]['ShortTermOtherPayables'], num2_format)						
		worksheet_1.write(offset + 31, k+1, balance_sheet_sub_list[k]['ShortTermAdvancesCustomers'], num2_format)						
		worksheet_1.write(offset + 32, k+1, balance_sheet_sub_list[k]['ShortTermWithholdings'], num2_format)						
		worksheet_1.write(offset + 33, k+1, balance_sheet_sub_list[k]['ShortTermBorrowings'], num2_format)				
		worksheet_1.write(offset + 34, k+1, balance_sheet_sub_list[k]['CurrentPortionOfLongtermBorrowings'], num2_format)						
		worksheet_1.write(offset + 35, k+1, balance_sheet_sub_list[k]['CurrentTaxLiabilities'], num2_format)						
		worksheet_1.write(offset + 36, k+1, balance_sheet_sub_list[k]['OtherCurrentFinancialLiabilities'], num2_format)						
		worksheet_1.write(offset + 37, k+1, balance_sheet_sub_list[k]['CurrentProvisions'], num2_format)						
		worksheet_1.write(offset + 38, k+1, balance_sheet_sub_list[k]['OtherCurrentLiabilities'], num2_format)						
		worksheet_1.write(offset + 39, k+1, balance_sheet_sub_list[k]['LiabilitiesIncludedInDisposal'], num2_format)						
		
		worksheet_1.write(offset + 40, k+1, balance_sheet_sub_list[k]['NoncurrentLiabilities'], num2_format)				
		worksheet_1.write(offset + 41, k+1, balance_sheet_sub_list[k]['LongTermTradeAndOtherNonCurrent'], num2_format)						
		worksheet_1.write(offset + 42, k+1, balance_sheet_sub_list[k]['BondsIssued'], num2_format)						
		worksheet_1.write(offset + 43, k+1, balance_sheet_sub_list[k]['LongTermBorrowingsGross'], num2_format)			
		worksheet_1.write(offset + 44, k+1, balance_sheet_sub_list[k]['OtherNoncurrentFinancial'], num2_format)						
		worksheet_1.write(offset + 45, k+1, balance_sheet_sub_list[k]['NoncurrentProvisions'], num2_format)						
		worksheet_1.write(offset + 46, k+1, balance_sheet_sub_list[k]['PostemploymentBenefitObligations'], num2_format)						
		worksheet_1.write(offset + 47, k+1, balance_sheet_sub_list[k]['DeferredTaxLiabilities'], num2_format)			
		worksheet_1.write(offset + 48, k+1, balance_sheet_sub_list[k]['OtherNonCurrentLiabilities'], num2_format)						
		
		worksheet_1.write(offset + 49, k+1, balance_sheet_sub_list[k]['Liabilities'], num2_format)						
		worksheet_1.write(offset + 50, k+1, balance_sheet_sub_list[k]['IssuedCapital'], num2_format)						
		worksheet_1.write(offset + 51, k+1, balance_sheet_sub_list[k]['SharePremium'], num2_format)						
		worksheet_1.write(offset + 52, k+1, balance_sheet_sub_list[k]['RetainedEarnings'], num2_format)					
		worksheet_1.write(offset + 53, k+1, balance_sheet_sub_list[k]['Equity'], num2_format)							
	
	#worksheet_6 = workbook.add_worksheet('연결_재무상태표_분기')
	
	worksheet_2.write(offset + 0, 0, "연결 재무상태표", filter_format2)
	worksheet_2.set_column('A:A', 30)
	worksheet_2.write(offset + 1, 0, "유동자산")
	worksheet_2.write(offset + 2, 0, "  현금및현금성자산")
	worksheet_2.write(offset + 3, 0, "  단기금융상품")
	worksheet_2.write(offset + 4, 0, "  기타유동금융자산")
	worksheet_2.write(offset + 5, 0, "  매출채권")
	worksheet_2.write(offset + 6, 0, "  기타유동채권")
	worksheet_2.write(offset + 7, 0, "  단기미수금")
	worksheet_2.write(offset + 8, 0, "  단기선급금")
	worksheet_2.write(offset + 9, 0, "  단기선급비용")
	worksheet_2.write(offset + 10, 0, "  재고자산")
	worksheet_2.write(offset + 11, 0, "  기타유동자산")
	worksheet_2.write(offset + 12, 0, "  당기법인세자산")
	worksheet_2.write(offset + 13, 0, "  매각예정분류자산")

	worksheet_2.write(offset + 14, 0, "비유동자산")
	worksheet_2.write(offset + 15, 0, "  장기금융상품")
	worksheet_2.write(offset + 16, 0, "  기타비유동금융자산")
	worksheet_2.write(offset + 17, 0, "  장기매출채권 및 기타비유동채권")
	worksheet_2.write(offset + 18, 0, "  장기매출채권")
	worksheet_2.write(offset + 19, 0, "  유형자산")
	worksheet_2.write(offset + 20, 0, "  투자부동산")
	worksheet_2.write(offset + 21, 0, "  영업권")
	worksheet_2.write(offset + 22, 0, "  영업권 이외의 무형자산")
	worksheet_2.write(offset + 23, 0, "  지분법적용 투자지분")
	worksheet_2.write(offset + 24, 0, "  이연법인세자산")
	worksheet_2.write(offset + 25, 0, "  기타비유동자산")
	worksheet_2.write(offset + 26, 0, "자산총계", filter_format2)

	worksheet_2.write(offset + 27, 0, "유동부채")
	worksheet_2.write(offset + 28, 0, "  매입채무 및 기타유동채무")
	worksheet_2.write(offset + 29, 0, "  단기매입채무")
	worksheet_2.write(offset + 30, 0, "  단기미지급금")
	worksheet_2.write(offset + 31, 0, "  단기선수금")
	worksheet_2.write(offset + 32, 0, "  단기예수금")
	worksheet_2.write(offset + 33, 0, "  단기차입금")
	worksheet_2.write(offset + 34, 0, "  유동성장기차입금")
	worksheet_2.write(offset + 35, 0, "  당기법인세부채")
	worksheet_2.write(offset + 36, 0, "  기타유동금융부채")
	worksheet_2.write(offset + 37, 0, "  유동충당부채")
	worksheet_2.write(offset + 38, 0, "  기타유동부채")
	worksheet_2.write(offset + 39, 0, "  매각예정분류부채")
	
	worksheet_2.write(offset + 40, 0, "비유동부채")
	worksheet_2.write(offset + 41, 0, "  장기매입채무 및 기타비유동채무")
	worksheet_2.write(offset + 42, 0, "  사채")
	worksheet_2.write(offset + 43, 0, "  장기차입금")
	worksheet_2.write(offset + 44, 0, "  기타비유동금융부채")
	worksheet_2.write(offset + 45, 0, "  비유동충당부채")
	worksheet_2.write(offset + 46, 0, "  퇴직급여부채")
	worksheet_2.write(offset + 47, 0, "  이연법인세부채")
	worksheet_2.write(offset + 48, 0, "  기타비유동부채")
	worksheet_2.write(offset + 49, 0, "부채총계", filter_format2)
	
	worksheet_2.write(offset + 50, 0, "  자본금")
	worksheet_2.write(offset + 51, 0, "  자본잉여금")
	worksheet_2.write(offset + 52, 0, "  이익잉여금")
	worksheet_2.write(offset + 53, 0, "자본총계", filter_format2)
	
	worksheet_2.write(offset + 0, 1,	"2017.4Q", filter_format)
	worksheet_2.write(offset + 0, 2,	"2017.3Q", filter_format)
	worksheet_2.write(offset + 0, 3, 	"2017.2Q", filter_format)
	worksheet_2.write(offset + 0, 4, 	"2017.1Q", filter_format)
	worksheet_2.write(offset + 0, 5, 	"2016.4Q", filter_format)
	worksheet_2.write(offset + 0, 6, 	"2016.3Q", filter_format)
	worksheet_2.write(offset + 0, 7, 	"2016.2Q", filter_format)
	worksheet_2.write(offset + 0, 8, 	"2016.1Q", filter_format)
	
	for k in range(len(balance_sheet_sub_list_q)):
	
		worksheet_2.write(offset + 1, k+1, balance_sheet_sub_list_q[k]['CurrentAssets'], num2_format)						
		worksheet_2.write(offset + 2, k+1, balance_sheet_sub_list_q[k]['CashAndCashEquivalents'], num2_format)			
		worksheet_2.write(offset + 3, k+1, balance_sheet_sub_list_q[k]['ShortTermDeposits'], num2_format)					
		worksheet_2.write(offset + 4, k+1, balance_sheet_sub_list_q[k]['OtherCurrentFinancialAssets'], num2_format)	
		worksheet_2.write(offset + 5, k+1, balance_sheet_sub_list_q[k]['ShortTermTradeReceivable'], num2_format)	
		worksheet_2.write(offset + 6, k+1, balance_sheet_sub_list_q[k]['TradeAndOtherCurrentReceivables'], num2_format)	
		worksheet_2.write(offset + 7, k+1, balance_sheet_sub_list_q[k]['ShortTermOtherReceivables'], num2_format)	
		worksheet_2.write(offset + 8, k+1, balance_sheet_sub_list_q[k]['ShortTermAdvancePayments'], num2_format)	
		worksheet_2.write(offset + 9, k+1, balance_sheet_sub_list_q[k]['ShortTermPrepaidExpenses'], num2_format)	
		worksheet_2.write(offset + 10, k+1, balance_sheet_sub_list_q[k]['Inventories'], num2_format)						
		worksheet_2.write(offset + 11, k+1, balance_sheet_sub_list_q[k]['OtherCurrentNonfinancialAssets'], num2_format)	
		worksheet_2.write(offset + 12, k+1, balance_sheet_sub_list_q[k]['CurrentTaxAssets'], num2_format)	
		worksheet_2.write(offset + 13, k+1, balance_sheet_sub_list_q[k]['NoncurrentAssetsOrDisposal'], num2_format)	
		
		worksheet_2.write(offset + 14, k+1, balance_sheet_sub_list_q[k]['NoncurrentAssets'], num2_format)					
		worksheet_2.write(offset + 15, k+1, balance_sheet_sub_list_q[k]['LongTermDeposits'], num2_format)						
		worksheet_2.write(offset + 16, k+1, balance_sheet_sub_list_q[k]['OtherNoncurrentFinancialAssets'], num2_format)						
		worksheet_2.write(offset + 17, k+1, balance_sheet_sub_list_q[k]['LongTermTradeAndOther'], num2_format)						
		worksheet_2.write(offset + 18, k+1, balance_sheet_sub_list_q[k]['LongTermTradeReceivablesGross'], num2_format)						
		worksheet_2.write(offset + 19, k+1, balance_sheet_sub_list_q[k]['PropertyPlantAndEquipment'], num2_format)			
		worksheet_2.write(offset + 20, k+1, balance_sheet_sub_list_q[k]['InvestmentProperty'], num2_format)				
		worksheet_2.write(offset + 21, k+1, balance_sheet_sub_list_q[k]['GoodwillGross'], num2_format)						
		worksheet_2.write(offset + 22, k+1, balance_sheet_sub_list_q[k]['IntangibleAssetsOtherThanGoodwill'], num2_format)						
		worksheet_2.write(offset + 23, k+1, balance_sheet_sub_list_q[k]['InvestmentAccounted'], num2_format)						
		worksheet_2.write(offset + 24, k+1, balance_sheet_sub_list_q[k]['DeferredTaxAssets'], num2_format)					
		worksheet_2.write(offset + 25, k+1, balance_sheet_sub_list_q[k]['OtherNonCurrentAssets'], num2_format)						
		worksheet_2.write(offset + 26, k+1, balance_sheet_sub_list_q[k]['Assets'], num2_format)							
		
		worksheet_2.write(offset + 27, k+1, balance_sheet_sub_list_q[k]['CurrentLiabilities'], num2_format)				
		worksheet_2.write(offset + 28, k+1, balance_sheet_sub_list_q[k]['TradeAndOtherCurrentPayables'], num2_format)		
		worksheet_2.write(offset + 29, k+1, balance_sheet_sub_list_q[k]['ShortTermTradePayables'], num2_format)						
		worksheet_2.write(offset + 30, k+1, balance_sheet_sub_list_q[k]['ShortTermOtherPayables'], num2_format)						
		worksheet_2.write(offset + 31, k+1, balance_sheet_sub_list_q[k]['ShortTermAdvancesCustomers'], num2_format)						
		worksheet_2.write(offset + 32, k+1, balance_sheet_sub_list_q[k]['ShortTermWithholdings'], num2_format)						
		worksheet_2.write(offset + 33, k+1, balance_sheet_sub_list_q[k]['ShortTermBorrowings'], num2_format)				
		worksheet_2.write(offset + 34, k+1, balance_sheet_sub_list_q[k]['CurrentPortionOfLongtermBorrowings'], num2_format)						
		worksheet_2.write(offset + 35, k+1, balance_sheet_sub_list_q[k]['CurrentTaxLiabilities'], num2_format)						
		worksheet_2.write(offset + 36, k+1, balance_sheet_sub_list_q[k]['OtherCurrentFinancialLiabilities'], num2_format)						
		worksheet_2.write(offset + 37, k+1, balance_sheet_sub_list_q[k]['CurrentProvisions'], num2_format)						
		worksheet_2.write(offset + 38, k+1, balance_sheet_sub_list_q[k]['OtherCurrentLiabilities'], num2_format)						
		worksheet_2.write(offset + 39, k+1, balance_sheet_sub_list_q[k]['LiabilitiesIncludedInDisposal'], num2_format)						
		
		worksheet_2.write(offset + 40, k+1, balance_sheet_sub_list_q[k]['NoncurrentLiabilities'], num2_format)				
		worksheet_2.write(offset + 41, k+1, balance_sheet_sub_list_q[k]['LongTermTradeAndOtherNonCurrent'], num2_format)						
		worksheet_2.write(offset + 42, k+1, balance_sheet_sub_list_q[k]['BondsIssued'], num2_format)						
		worksheet_2.write(offset + 43, k+1, balance_sheet_sub_list_q[k]['LongTermBorrowingsGross'], num2_format)			
		worksheet_2.write(offset + 44, k+1, balance_sheet_sub_list_q[k]['OtherNoncurrentFinancial'], num2_format)						
		worksheet_2.write(offset + 45, k+1, balance_sheet_sub_list_q[k]['NoncurrentProvisions'], num2_format)						
		worksheet_2.write(offset + 46, k+1, balance_sheet_sub_list_q[k]['PostemploymentBenefitObligations'], num2_format)						
		worksheet_2.write(offset + 47, k+1, balance_sheet_sub_list_q[k]['DeferredTaxLiabilities'], num2_format)			
		worksheet_2.write(offset + 48, k+1, balance_sheet_sub_list_q[k]['OtherNonCurrentLiabilities'], num2_format)						
		
		worksheet_2.write(offset + 49, k+1, balance_sheet_sub_list_q[k]['Liabilities'], num2_format)						
		worksheet_2.write(offset + 50, k+1, balance_sheet_sub_list_q[k]['IssuedCapital'], num2_format)						
		worksheet_2.write(offset + 51, k+1, balance_sheet_sub_list_q[k]['SharePremium'], num2_format)						
		worksheet_2.write(offset + 52, k+1, balance_sheet_sub_list_q[k]['RetainedEarnings'], num2_format)					
		worksheet_2.write(offset + 53, k+1, balance_sheet_sub_list_q[k]['Equity'], num2_format)							
	
	#worksheet_7 = workbook.add_worksheet('개별_재무상태표_year')
	
	worksheet_3.write(offset + 0, 0, "개별 재무상태표", filter_format2)
	worksheet_3.set_column('A:A', 30)
	worksheet_3.write(offset + 1, 0, "유동자산")
	worksheet_3.write(offset + 2, 0, "  현금및현금성자산")
	worksheet_3.write(offset + 3, 0, "  단기금융상품")
	worksheet_3.write(offset + 4, 0, "  기타유동금융자산")
	worksheet_3.write(offset + 5, 0, "  매출채권")
	worksheet_3.write(offset + 6, 0, "  기타유동채권")
	worksheet_3.write(offset + 7, 0, "  단기미수금")
	worksheet_3.write(offset + 8, 0, "  단기선급금")
	worksheet_3.write(offset + 9, 0, "  단기선급비용")
	worksheet_3.write(offset + 10, 0, "  재고자산")
	worksheet_3.write(offset + 11, 0, "  기타유동자산")
	worksheet_3.write(offset + 12, 0, "  당기법인세자산")
	worksheet_3.write(offset + 13, 0, "  매각예정분류자산")

	worksheet_3.write(offset + 14, 0, "비유동자산")
	worksheet_3.write(offset + 15, 0, "  장기금융상품")
	worksheet_3.write(offset + 16, 0, "  기타비유동금융자산")
	worksheet_3.write(offset + 17, 0, "  장기매출채권 및 기타비유동채권")
	worksheet_3.write(offset + 18, 0, "  장기매출채권")
	worksheet_3.write(offset + 19, 0, "  유형자산")
	worksheet_3.write(offset + 20, 0, "  투자부동산")
	worksheet_3.write(offset + 21, 0, "  영업권")
	worksheet_3.write(offset + 22, 0, "  영업권 이외의 무형자산")
	worksheet_3.write(offset + 23, 0, "  지분법적용 투자지분")
	worksheet_3.write(offset + 24, 0, "  이연법인세자산")
	worksheet_3.write(offset + 25, 0, "  기타비유동자산")
	worksheet_3.write(offset + 26, 0, "자산총계", filter_format2)

	worksheet_3.write(offset + 27, 0, "유동부채")
	worksheet_3.write(offset + 28, 0, "  매입채무 및 기타유동채무")
	worksheet_3.write(offset + 29, 0, "  단기매입채무")
	worksheet_3.write(offset + 30, 0, "  단기미지급금")
	worksheet_3.write(offset + 31, 0, "  단기선수금")
	worksheet_3.write(offset + 32, 0, "  단기예수금")
	worksheet_3.write(offset + 33, 0, "  단기차입금")
	worksheet_3.write(offset + 34, 0, "  유동성장기차입금")
	worksheet_3.write(offset + 35, 0, "  당기법인세부채")
	worksheet_3.write(offset + 36, 0, "  기타유동금융부채")
	worksheet_3.write(offset + 37, 0, "  유동충당부채")
	worksheet_3.write(offset + 38, 0, "  기타유동부채")
	worksheet_3.write(offset + 39, 0, "  매각예정분류부채")
	
	worksheet_3.write(offset + 40, 0, "비유동부채")
	worksheet_3.write(offset + 41, 0, "  장기매입채무 및 기타비유동채무")
	worksheet_3.write(offset + 42, 0, "  사채")
	worksheet_3.write(offset + 43, 0, "  장기차입금")
	worksheet_3.write(offset + 44, 0, "  기타비유동금융부채")
	worksheet_3.write(offset + 45, 0, "  비유동충당부채")
	worksheet_3.write(offset + 46, 0, "  퇴직급여부채")
	worksheet_3.write(offset + 47, 0, "  이연법인세부채")
	worksheet_3.write(offset + 48, 0, "  기타비유동부채")
	worksheet_3.write(offset + 49, 0, "부채총계", filter_format2)
	
	worksheet_3.write(offset + 50, 0, "  자본금")
	worksheet_3.write(offset + 51, 0, "  자본잉여금")
	worksheet_3.write(offset + 52, 0, "  이익잉여금")
	worksheet_3.write(offset + 53, 0, "자본총계", filter_format2)
	
	worksheet_3.write(offset + 0, 1,	"2017년", filter_format)
	worksheet_3.write(offset + 0, 2, 	"2016년", filter_format)
	worksheet_3.write(offset + 0, 3, 	"2015년", filter_format)
	worksheet_3.write(offset + 0, 4, 	"2014년", filter_format)
	worksheet_3.write(offset + 0, 5,	"2013년", filter_format)
	
	for k in range(len(individual_balance_sheet_sub_list)):
	
		worksheet_3.write(offset + 1, k+1, individual_balance_sheet_sub_list[k]['CurrentAssets'], num2_format)						
		worksheet_3.write(offset + 2, k+1, individual_balance_sheet_sub_list[k]['CashAndCashEquivalents'], num2_format)			
		worksheet_3.write(offset + 3, k+1, individual_balance_sheet_sub_list[k]['ShortTermDeposits'], num2_format)					
		worksheet_3.write(offset + 4, k+1, individual_balance_sheet_sub_list[k]['OtherCurrentFinancialAssets'], num2_format)	
		worksheet_3.write(offset + 5, k+1, individual_balance_sheet_sub_list[k]['ShortTermTradeReceivable'], num2_format)	
		worksheet_3.write(offset + 6, k+1, individual_balance_sheet_sub_list[k]['TradeAndOtherCurrentReceivables'], num2_format)	
		worksheet_3.write(offset + 7, k+1, individual_balance_sheet_sub_list[k]['ShortTermOtherReceivables'], num2_format)	
		worksheet_3.write(offset + 8, k+1, individual_balance_sheet_sub_list[k]['ShortTermAdvancePayments'], num2_format)	
		worksheet_3.write(offset + 9, k+1, individual_balance_sheet_sub_list[k]['ShortTermPrepaidExpenses'], num2_format)	
		worksheet_3.write(offset + 10, k+1, individual_balance_sheet_sub_list[k]['Inventories'], num2_format)						
		worksheet_3.write(offset + 11, k+1, individual_balance_sheet_sub_list[k]['OtherCurrentNonfinancialAssets'], num2_format)	
		worksheet_3.write(offset + 12, k+1, individual_balance_sheet_sub_list[k]['CurrentTaxAssets'], num2_format)	
		worksheet_3.write(offset + 13, k+1, individual_balance_sheet_sub_list[k]['NoncurrentAssetsOrDisposal'], num2_format)	
		
		
		worksheet_3.write(offset + 14, k+1, individual_balance_sheet_sub_list[k]['NoncurrentAssets'], num2_format)					
		worksheet_3.write(offset + 15, k+1, individual_balance_sheet_sub_list[k]['LongTermDeposits'], num2_format)						
		worksheet_3.write(offset + 16, k+1, individual_balance_sheet_sub_list[k]['OtherNoncurrentFinancialAssets'], num2_format)						
		worksheet_3.write(offset + 17, k+1, individual_balance_sheet_sub_list[k]['LongTermTradeAndOther'], num2_format)						
		worksheet_3.write(offset + 18, k+1, individual_balance_sheet_sub_list[k]['LongTermTradeReceivablesGross'], num2_format)						
		worksheet_3.write(offset + 19, k+1, individual_balance_sheet_sub_list[k]['PropertyPlantAndEquipment'], num2_format)			
		worksheet_3.write(offset + 20, k+1, individual_balance_sheet_sub_list[k]['InvestmentProperty'], num2_format)				
		worksheet_3.write(offset + 21, k+1, individual_balance_sheet_sub_list[k]['GoodwillGross'], num2_format)						
		worksheet_3.write(offset + 22, k+1, individual_balance_sheet_sub_list[k]['IntangibleAssetsOtherThanGoodwill'], num2_format)						
		worksheet_3.write(offset + 23, k+1, individual_balance_sheet_sub_list[k]['InvestmentAccounted'], num2_format)						
		worksheet_3.write(offset + 24, k+1, individual_balance_sheet_sub_list[k]['DeferredTaxAssets'], num2_format)					
		worksheet_3.write(offset + 25, k+1, individual_balance_sheet_sub_list[k]['OtherNonCurrentAssets'], num2_format)						
		worksheet_3.write(offset + 26, k+1, individual_balance_sheet_sub_list[k]['Assets'], num2_format)							
		
		
		worksheet_3.write(offset + 27, k+1, individual_balance_sheet_sub_list[k]['CurrentLiabilities'], num2_format)				
		worksheet_3.write(offset + 28, k+1, individual_balance_sheet_sub_list[k]['TradeAndOtherCurrentPayables'], num2_format)		
		worksheet_3.write(offset + 29, k+1, individual_balance_sheet_sub_list[k]['ShortTermTradePayables'], num2_format)						
		worksheet_3.write(offset + 30, k+1, individual_balance_sheet_sub_list[k]['ShortTermOtherPayables'], num2_format)						
		worksheet_3.write(offset + 31, k+1, individual_balance_sheet_sub_list[k]['ShortTermAdvancesCustomers'], num2_format)						
		worksheet_3.write(offset + 32, k+1, individual_balance_sheet_sub_list[k]['ShortTermWithholdings'], num2_format)						
		worksheet_3.write(offset + 33, k+1, individual_balance_sheet_sub_list[k]['ShortTermBorrowings'], num2_format)				
		worksheet_3.write(offset + 34, k+1, individual_balance_sheet_sub_list[k]['CurrentPortionOfLongtermBorrowings'], num2_format)						
		worksheet_3.write(offset + 35, k+1, individual_balance_sheet_sub_list[k]['CurrentTaxLiabilities'], num2_format)						
		worksheet_3.write(offset + 36, k+1, individual_balance_sheet_sub_list[k]['OtherCurrentFinancialLiabilities'], num2_format)						
		worksheet_3.write(offset + 37, k+1, individual_balance_sheet_sub_list[k]['CurrentProvisions'], num2_format)						
		worksheet_3.write(offset + 38, k+1, individual_balance_sheet_sub_list[k]['OtherCurrentLiabilities'], num2_format)						
		worksheet_3.write(offset + 39, k+1, individual_balance_sheet_sub_list[k]['LiabilitiesIncludedInDisposal'], num2_format)						
		
		
		
		worksheet_3.write(offset + 40, k+1, individual_balance_sheet_sub_list[k]['NoncurrentLiabilities'], num2_format)				
		worksheet_3.write(offset + 41, k+1, individual_balance_sheet_sub_list[k]['LongTermTradeAndOtherNonCurrent'], num2_format)						
		worksheet_3.write(offset + 42, k+1, individual_balance_sheet_sub_list[k]['BondsIssued'], num2_format)						
		worksheet_3.write(offset + 43, k+1, individual_balance_sheet_sub_list[k]['LongTermBorrowingsGross'], num2_format)			
		worksheet_3.write(offset + 44, k+1, individual_balance_sheet_sub_list[k]['OtherNoncurrentFinancial'], num2_format)						
		worksheet_3.write(offset + 45, k+1, individual_balance_sheet_sub_list[k]['NoncurrentProvisions'], num2_format)						
		worksheet_3.write(offset + 46, k+1, individual_balance_sheet_sub_list[k]['PostemploymentBenefitObligations'], num2_format)						
		worksheet_3.write(offset + 47, k+1, individual_balance_sheet_sub_list[k]['DeferredTaxLiabilities'], num2_format)			
		worksheet_3.write(offset + 48, k+1, individual_balance_sheet_sub_list[k]['OtherNonCurrentLiabilities'], num2_format)						
		
		worksheet_3.write(offset + 49, k+1, individual_balance_sheet_sub_list[k]['Liabilities'], num2_format)						
		worksheet_3.write(offset + 50, k+1, individual_balance_sheet_sub_list[k]['IssuedCapital'], num2_format)						
		worksheet_3.write(offset + 51, k+1, individual_balance_sheet_sub_list[k]['SharePremium'], num2_format)						
		worksheet_3.write(offset + 52, k+1, individual_balance_sheet_sub_list[k]['RetainedEarnings'], num2_format)					
		worksheet_3.write(offset + 53, k+1, individual_balance_sheet_sub_list[k]['Equity'], num2_format)							
	
	#worksheet_8 = workbook.add_worksheet('개별_재무상태표_분기')
	
	worksheet_4.write(offset + 0, 0, "개별 재무상태표", filter_format2)
	worksheet_4.set_column('A:A', 30)
	worksheet_4.write(offset + 1, 0, "유동자산")
	worksheet_4.write(offset + 2, 0, "  현금및현금성자산")
	worksheet_4.write(offset + 3, 0, "  단기금융상품")
	worksheet_4.write(offset + 4, 0, "  기타유동금융자산")
	worksheet_4.write(offset + 5, 0, "  매출채권")
	worksheet_4.write(offset + 6, 0, "  기타유동채권")
	worksheet_4.write(offset + 7, 0, "  단기미수금")
	worksheet_4.write(offset + 8, 0, "  단기선급금")
	worksheet_4.write(offset + 9, 0, "  단기선급비용")
	worksheet_4.write(offset + 10, 0, "  재고자산")
	worksheet_4.write(offset + 11, 0, "  기타유동자산")
	worksheet_4.write(offset + 12, 0, "  당기법인세자산")
	worksheet_4.write(offset + 13, 0, "  매각예정분류자산")

	worksheet_4.write(offset + 14, 0, "비유동자산")
	worksheet_4.write(offset + 15, 0, "  장기금융상품")
	worksheet_4.write(offset + 16, 0, "  기타비유동금융자산")
	worksheet_4.write(offset + 17, 0, "  장기매출채권 및 기타비유동채권")
	worksheet_4.write(offset + 18, 0, "  장기매출채권")
	worksheet_4.write(offset + 19, 0, "  유형자산")
	worksheet_4.write(offset + 20, 0, "  투자부동산")
	worksheet_4.write(offset + 21, 0, "  영업권")
	worksheet_4.write(offset + 22, 0, "  영업권 이외의 무형자산")
	worksheet_4.write(offset + 23, 0, "  지분법적용 투자지분")
	worksheet_4.write(offset + 24, 0, "  이연법인세자산")
	worksheet_4.write(offset + 25, 0, "  기타비유동자산")
	worksheet_4.write(offset + 26, 0, "자산총계", filter_format2)

	worksheet_4.write(offset + 27, 0, "유동부채")
	worksheet_4.write(offset + 28, 0, "  매입채무 및 기타유동채무")
	worksheet_4.write(offset + 29, 0, "  단기매입채무")
	worksheet_4.write(offset + 30, 0, "  단기미지급금")
	worksheet_4.write(offset + 31, 0, "  단기선수금")
	worksheet_4.write(offset + 32, 0, "  단기예수금")
	worksheet_4.write(offset + 33, 0, "  단기차입금")
	worksheet_4.write(offset + 34, 0, "  유동성장기차입금")
	worksheet_4.write(offset + 35, 0, "  당기법인세부채")
	worksheet_4.write(offset + 36, 0, "  기타유동금융부채")
	worksheet_4.write(offset + 37, 0, "  유동충당부채")
	worksheet_4.write(offset + 38, 0, "  기타유동부채")
	worksheet_4.write(offset + 39, 0, "  매각예정분류부채")
	
	worksheet_4.write(offset + 40, 0, "비유동부채")
	worksheet_4.write(offset + 41, 0, "  장기매입채무 및 기타비유동채무")
	worksheet_4.write(offset + 42, 0, "  사채")
	worksheet_4.write(offset + 43, 0, "  장기차입금")
	worksheet_4.write(offset + 44, 0, "  기타비유동금융부채")
	worksheet_4.write(offset + 45, 0, "  비유동충당부채")
	worksheet_4.write(offset + 46, 0, "  퇴직급여부채")
	worksheet_4.write(offset + 47, 0, "  이연법인세부채")
	worksheet_4.write(offset + 48, 0, "  기타비유동부채")
	worksheet_4.write(offset + 49, 0, "부채총계", filter_format2)
	
	worksheet_4.write(offset + 50, 0, "  자본금")
	worksheet_4.write(offset + 51, 0, "  자본잉여금")
	worksheet_4.write(offset + 52, 0, "  이익잉여금")
	worksheet_4.write(offset + 53, 0, "자본총계", filter_format2)
	
	worksheet_4.write(offset + 0, 1,	"2017.4Q", filter_format)
	worksheet_4.write(offset + 0, 2,	"2017.3Q", filter_format)
	worksheet_4.write(offset + 0, 3, 	"2017.2Q", filter_format)
	worksheet_4.write(offset + 0, 4, 	"2017.1Q", filter_format)
	worksheet_4.write(offset + 0, 5, 	"2016.4Q", filter_format)
	worksheet_4.write(offset + 0, 6, 	"2016.3Q", filter_format)
	worksheet_4.write(offset + 0, 7, 	"2016.2Q", filter_format)
	worksheet_4.write(offset + 0, 8, 	"2016.1Q", filter_format)
	
	
	for k in range(len(individual_balance_sheet_sub_list_q)):
	
		worksheet_4.write(offset + 1, k+1, individual_balance_sheet_sub_list_q[k]['CurrentAssets'], num2_format)						
		worksheet_4.write(offset + 2, k+1, individual_balance_sheet_sub_list_q[k]['CashAndCashEquivalents'], num2_format)			
		worksheet_4.write(offset + 3, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermDeposits'], num2_format)					
		worksheet_4.write(offset + 4, k+1, individual_balance_sheet_sub_list_q[k]['OtherCurrentFinancialAssets'], num2_format)	
		worksheet_4.write(offset + 5, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermTradeReceivable'], num2_format)	
		worksheet_4.write(offset + 6, k+1, individual_balance_sheet_sub_list_q[k]['TradeAndOtherCurrentReceivables'], num2_format)	
		worksheet_4.write(offset + 7, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermOtherReceivables'], num2_format)	
		worksheet_4.write(offset + 8, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermAdvancePayments'], num2_format)	
		worksheet_4.write(offset + 9, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermPrepaidExpenses'], num2_format)	
		worksheet_4.write(offset + 10, k+1, individual_balance_sheet_sub_list_q[k]['Inventories'], num2_format)						
		worksheet_4.write(offset + 11, k+1, individual_balance_sheet_sub_list_q[k]['OtherCurrentNonfinancialAssets'], num2_format)	
		worksheet_4.write(offset + 12, k+1, individual_balance_sheet_sub_list_q[k]['CurrentTaxAssets'], num2_format)	
		worksheet_4.write(offset + 13, k+1, individual_balance_sheet_sub_list_q[k]['NoncurrentAssetsOrDisposal'], num2_format)	
		
		worksheet_4.write(offset + 14, k+1, individual_balance_sheet_sub_list_q[k]['NoncurrentAssets'], num2_format)					
		worksheet_4.write(offset + 15, k+1, individual_balance_sheet_sub_list_q[k]['LongTermDeposits'], num2_format)						
		worksheet_4.write(offset + 16, k+1, individual_balance_sheet_sub_list_q[k]['OtherNoncurrentFinancialAssets'], num2_format)						
		worksheet_4.write(offset + 17, k+1, individual_balance_sheet_sub_list_q[k]['LongTermTradeAndOther'], num2_format)						
		worksheet_4.write(offset + 18, k+1, individual_balance_sheet_sub_list_q[k]['LongTermTradeReceivablesGross'], num2_format)						
		worksheet_4.write(offset + 19, k+1, individual_balance_sheet_sub_list_q[k]['PropertyPlantAndEquipment'], num2_format)			
		worksheet_4.write(offset + 20, k+1, individual_balance_sheet_sub_list_q[k]['InvestmentProperty'], num2_format)				
		worksheet_4.write(offset + 21, k+1, individual_balance_sheet_sub_list_q[k]['GoodwillGross'], num2_format)						
		worksheet_4.write(offset + 22, k+1, individual_balance_sheet_sub_list_q[k]['IntangibleAssetsOtherThanGoodwill'], num2_format)						
		worksheet_4.write(offset + 23, k+1, individual_balance_sheet_sub_list_q[k]['InvestmentAccounted'], num2_format)						
		worksheet_4.write(offset + 24, k+1, individual_balance_sheet_sub_list_q[k]['DeferredTaxAssets'], num2_format)					
		worksheet_4.write(offset + 25, k+1, individual_balance_sheet_sub_list_q[k]['OtherNonCurrentAssets'], num2_format)						
		worksheet_4.write(offset + 26, k+1, individual_balance_sheet_sub_list_q[k]['Assets'], num2_format)							
		
		worksheet_4.write(offset + 27, k+1, individual_balance_sheet_sub_list_q[k]['CurrentLiabilities'], num2_format)				
		worksheet_4.write(offset + 28, k+1, individual_balance_sheet_sub_list_q[k]['TradeAndOtherCurrentPayables'], num2_format)		
		worksheet_4.write(offset + 29, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermTradePayables'], num2_format)						
		worksheet_4.write(offset + 30, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermOtherPayables'], num2_format)						
		worksheet_4.write(offset + 31, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermAdvancesCustomers'], num2_format)						
		worksheet_4.write(offset + 32, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermWithholdings'], num2_format)						
		worksheet_4.write(offset + 33, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermBorrowings'], num2_format)				
		worksheet_4.write(offset + 34, k+1, individual_balance_sheet_sub_list_q[k]['CurrentPortionOfLongtermBorrowings'], num2_format)						
		worksheet_4.write(offset + 35, k+1, individual_balance_sheet_sub_list_q[k]['CurrentTaxLiabilities'], num2_format)						
		worksheet_4.write(offset + 36, k+1, individual_balance_sheet_sub_list_q[k]['OtherCurrentFinancialLiabilities'], num2_format)						
		worksheet_4.write(offset + 37, k+1, individual_balance_sheet_sub_list_q[k]['CurrentProvisions'], num2_format)						
		worksheet_4.write(offset + 38, k+1, individual_balance_sheet_sub_list_q[k]['OtherCurrentLiabilities'], num2_format)						
		worksheet_4.write(offset + 39, k+1, individual_balance_sheet_sub_list_q[k]['LiabilitiesIncludedInDisposal'], num2_format)						
		
		worksheet_4.write(offset + 40, k+1, individual_balance_sheet_sub_list_q[k]['NoncurrentLiabilities'], num2_format)				
		worksheet_4.write(offset + 41, k+1, individual_balance_sheet_sub_list_q[k]['LongTermTradeAndOtherNonCurrent'], num2_format)						
		worksheet_4.write(offset + 42, k+1, individual_balance_sheet_sub_list_q[k]['BondsIssued'], num2_format)						
		worksheet_4.write(offset + 43, k+1, individual_balance_sheet_sub_list_q[k]['LongTermBorrowingsGross'], num2_format)			
		worksheet_4.write(offset + 44, k+1, individual_balance_sheet_sub_list_q[k]['OtherNoncurrentFinancial'], num2_format)						
		worksheet_4.write(offset + 45, k+1, individual_balance_sheet_sub_list_q[k]['NoncurrentProvisions'], num2_format)						
		worksheet_4.write(offset + 46, k+1, individual_balance_sheet_sub_list_q[k]['PostemploymentBenefitObligations'], num2_format)						
		worksheet_4.write(offset + 47, k+1, individual_balance_sheet_sub_list_q[k]['DeferredTaxLiabilities'], num2_format)			
		worksheet_4.write(offset + 48, k+1, individual_balance_sheet_sub_list_q[k]['OtherNonCurrentLiabilities'], num2_format)						
		
		worksheet_4.write(offset + 49, k+1, individual_balance_sheet_sub_list_q[k]['Liabilities'], num2_format)						
		worksheet_4.write(offset + 50, k+1, individual_balance_sheet_sub_list_q[k]['IssuedCapital'], num2_format)						
		worksheet_4.write(offset + 51, k+1, individual_balance_sheet_sub_list_q[k]['SharePremium'], num2_format)						
		worksheet_4.write(offset + 52, k+1, individual_balance_sheet_sub_list_q[k]['RetainedEarnings'], num2_format)					
		worksheet_4.write(offset + 53, k+1, individual_balance_sheet_sub_list_q[k]['Equity'], num2_format)							
	
	#worksheet_9 = workbook.add_worksheet('연결_현금흐름표_year')

	offset = offset + len(balance_sheet_sub_list[0]) + 2

	worksheet_1.write(offset + 0, 0, "연결 현금흐름표", filter_format2)
	worksheet_1.set_column('A:A', 30)
	worksheet_1.write(offset + 1, 0, "영업활동현금흐름")
	worksheet_1.write(offset + 2, 0, " 당기순이익")
	worksheet_1.write(offset + 3, 0, " 당기순이익조정을 위한 가감")
	worksheet_1.write(offset + 4, 0, " 감가상각비")
	worksheet_1.write(offset + 5, 0, "투자활동현금흐름")
	worksheet_1.write(offset + 6, 0, " 유형자산의 취득")
	worksheet_1.write(offset + 7, 0, " 무형자산의 취득")
	worksheet_1.write(offset + 8, 0, " 투자부동산의 취득")
	worksheet_1.write(offset + 9, 0, " 유형자산의 처분")
	worksheet_1.write(offset + 10, 0, " 무형자산의 처분")
	worksheet_1.write(offset + 11, 0, " 투자부동산의 처분")
	worksheet_1.write(offset + 12, 0, "재무활동현금흐름")
	worksheet_1.write(offset + 13, 0, " 단기차입금의 증가")
	worksheet_1.write(offset + 14, 0, " 배당금의 지급")
	worksheet_1.write(offset + 15, 0, "기초현금및현금성자산")
	worksheet_1.write(offset + 16, 0, "기말현금및현금성자산")
	
	worksheet_1.write(offset + 0, 1,	"2017년", filter_format)
	worksheet_1.write(offset + 0, 2, 	"2016년", filter_format)
	worksheet_1.write(offset + 0, 3, 	"2015년", filter_format)
	worksheet_1.write(offset + 0, 4, 	"2014년", filter_format)
	worksheet_1.write(offset + 0, 5,	"2013년", filter_format)
	
	for k in range(len(cashflow_statement_sub_list)):
	
		worksheet_1.write(offset + 1, k+1,	cashflow_statement_sub_list[k]['CashFlowsFromUsedInOperatingActivities'], num2_format)
		worksheet_1.write(offset + 2, k+1, 	cashflow_statement_sub_list[k]['ProfitLossForStatementOfCashFlows'], num2_format)
		worksheet_1.write(offset + 3, k+1, 	cashflow_statement_sub_list[k]['AdjustmentsForReconcileProfitLoss'], num2_format)
		worksheet_1.write(offset + 4, k+1, 	cashflow_statement_sub_list[k]['AdjustmentsForDepreciationExpense'], num2_format)
		worksheet_1.write(offset + 5, k+1, 	cashflow_statement_sub_list[k]['CashFlowsFromUsedInInvestingActivities'], num2_format)
		worksheet_1.write(offset + 6, k+1, 	cashflow_statement_sub_list[k]['PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities'], num2_format)
		worksheet_1.write(offset + 7, k+1, 	cashflow_statement_sub_list[k]['PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities'], num2_format)
		worksheet_1.write(offset + 8, k+1, 	cashflow_statement_sub_list[k]['PurchaseOfInvestmentProperty'], num2_format)
		worksheet_1.write(offset + 9, k+1, 	cashflow_statement_sub_list[k]['ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities'], num2_format)
		worksheet_1.write(offset + 10, k+1, 	cashflow_statement_sub_list[k]['ProceedsFromSalesOfIntangibleAssetsClassifiedAsInvestingActivities'], num2_format)
		worksheet_1.write(offset + 11, k+1,	cashflow_statement_sub_list[k]['ProceedsFromSalesOfInvestmentProperty'], num2_format) 
		worksheet_1.write(offset + 12, k+1, 	cashflow_statement_sub_list[k]['CashFlowsFromUsedInFinancingActivities'], num2_format)
		worksheet_1.write(offset + 13, k+1, 	cashflow_statement_sub_list[k]['ProceedsFromShortTermBorrowings'], num2_format)	
		worksheet_1.write(offset + 14, k+1, 	cashflow_statement_sub_list[k]['DividendsPaidClassifiedAsFinancingActivities'], num2_format)
		worksheet_1.write(offset + 15, k+1, 	cashflow_statement_sub_list[k]['CashAndCashEquivalentsAtBeginningOfPeriodCf'], num2_format)
		worksheet_1.write(offset + 16, k+1, 	cashflow_statement_sub_list[k]['CashAndCashEquivalentsAtEndOfPeriodCf'], num2_format)

	#worksheet_10 = workbook.add_worksheet('연결_현금흐름표_분기')
	           
	worksheet_2.write(offset + 0, 0, "연결 현금흐름표", filter_format2)
	worksheet_2.set_column('A:A', 30)
	worksheet_2.write(offset + 1, 0, "영업활동현금흐름")
	worksheet_2.write(offset + 2, 0, " 당기순이익")
	worksheet_2.write(offset + 3, 0, " 당기순이익조정을 위한 가감")
	worksheet_2.write(offset + 4, 0, " 감가상각비")
	worksheet_2.write(offset + 5, 0, "투자활동현금흐름")
	worksheet_2.write(offset + 6, 0, " 유형자산의 취득")
	worksheet_2.write(offset + 7, 0, " 무형자산의 취득")
	worksheet_2.write(offset + 8, 0, " 투자부동산의 취득")
	worksheet_2.write(offset + 9, 0, " 유형자산의 처분")
	worksheet_2.write(offset + 10, 0, " 무형자산의 처분")
	worksheet_2.write(offset + 11, 0, " 투자부동산의 처분")
	worksheet_2.write(offset + 12, 0, "재무활동현금흐름")
	worksheet_2.write(offset + 13, 0, " 단기차입금의 증가")
	worksheet_2.write(offset + 14, 0, " 배당금의 지급")
	worksheet_2.write(offset + 15, 0, "기초현금및현금성자산")
	worksheet_2.write(offset + 16, 0, "기말현금및현금성자산")
	          
	worksheet_2.write(offset + 0, 1,	"2017.4Q", filter_format)
	worksheet_2.write(offset + 0, 2,	"2017.3Q", filter_format)
	worksheet_2.write(offset + 0, 3, 	"2017.2Q", filter_format)
	worksheet_2.write(offset + 0, 4, 	"2017.1Q", filter_format)
	worksheet_2.write(offset + 0, 5, 	"2016.4Q", filter_format)
	worksheet_2.write(offset + 0, 6, 	"2016.3Q", filter_format)
	worksheet_2.write(offset + 0, 7, 	"2016.2Q", filter_format)
	worksheet_2.write(offset + 0, 8, 	"2016.1Q", filter_format)
	
	
	for k in range(len(cashflow_statement_sub_list_q)):
	
		worksheet_2.write(offset + 1, k+1,	cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInOperatingActivities'], num2_format)
		worksheet_2.write(offset + 2, k+1, 	cashflow_statement_sub_list_q[k]['ProfitLossForStatementOfCashFlows'], num2_format)
		worksheet_2.write(offset + 3, k+1, 	cashflow_statement_sub_list_q[k]['AdjustmentsForReconcileProfitLoss'], num2_format)
		worksheet_2.write(offset + 4, k+1, 	cashflow_statement_sub_list_q[k]['AdjustmentsForDepreciationExpense'], num2_format)
		worksheet_2.write(offset + 5, k+1, 	cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInInvestingActivities'], num2_format)
		worksheet_2.write(offset + 6, k+1, 	cashflow_statement_sub_list_q[k]['PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities'], num2_format)
		worksheet_2.write(offset + 7, k+1, 	cashflow_statement_sub_list_q[k]['PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities'], num2_format)
		worksheet_2.write(offset + 8, k+1, 	cashflow_statement_sub_list_q[k]['PurchaseOfInvestmentProperty'], num2_format)
		worksheet_2.write(offset + 9, k+1,  cashflow_statement_sub_list_q[k]['ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities'], num2_format)
		worksheet_2.write(offset + 10, k+1, cashflow_statement_sub_list_q[k]['ProceedsFromSalesOfIntangibleAssetsClassifiedAsInvestingActivities'], num2_format)
		worksheet_2.write(offset + 11, k+1,	cashflow_statement_sub_list_q[k]['ProceedsFromSalesOfInvestmentProperty'], num2_format) 
		worksheet_2.write(offset + 12, k+1, cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInFinancingActivities'], num2_format)
		worksheet_2.write(offset + 13, k+1, cashflow_statement_sub_list_q[k]['ProceedsFromShortTermBorrowings'], num2_format)	
		worksheet_2.write(offset + 14, k+1, cashflow_statement_sub_list_q[k]['DividendsPaidClassifiedAsFinancingActivities'], num2_format)
		worksheet_2.write(offset + 15, k+1, cashflow_statement_sub_list_q[k]['CashAndCashEquivalentsAtBeginningOfPeriodCf'], num2_format)
		worksheet_2.write(offset + 16, k+1, cashflow_statement_sub_list_q[k]['CashAndCashEquivalentsAtEndOfPeriodCf'], num2_format)

	worksheet_3.write(offset + 0, 0, "연결 현금흐름표", filter_format2)
	worksheet_3.set_column('A:A', 30)
	worksheet_3.write(offset + 1, 0, "영업활동현금흐름")
	worksheet_3.write(offset + 2, 0, " 당기순이익")
	worksheet_3.write(offset + 3, 0, " 당기순이익조정을 위한 가감")
	worksheet_3.write(offset + 4, 0, " 감가상각비")
	worksheet_3.write(offset + 5, 0, "투자활동현금흐름")
	worksheet_3.write(offset + 6, 0, " 유형자산의 취득")
	worksheet_3.write(offset + 7, 0, " 무형자산의 취득")
	worksheet_3.write(offset + 8, 0, " 투자부동산의 취득")
	worksheet_3.write(offset + 9, 0, " 유형자산의 처분")
	worksheet_3.write(offset + 10, 0, " 무형자산의 처분")
	worksheet_3.write(offset + 11, 0, " 투자부동산의 처분")
	worksheet_3.write(offset + 12, 0, "재무활동현금흐름")
	worksheet_3.write(offset + 13, 0, " 단기차입금의 증가")
	worksheet_3.write(offset + 14, 0, " 배당금의 지급")
	worksheet_3.write(offset + 15, 0, "기초현금및현금성자산")
	worksheet_3.write(offset + 16, 0, "기말현금및현금성자산")
	
	worksheet_3.write(offset + 0, 1,	"2017년", filter_format)
	worksheet_3.write(offset + 0, 2, 	"2016년", filter_format)
	worksheet_3.write(offset + 0, 3, 	"2015년", filter_format)
	worksheet_3.write(offset + 0, 4, 	"2014년", filter_format)
	worksheet_3.write(offset + 0, 5,	"2013년", filter_format)
	
	for k in range(len(individual_cashflow_statement_sub_list)):
	
		worksheet_3.write(offset + 1, k+1,	individual_cashflow_statement_sub_list[k]['CashFlowsFromUsedInOperatingActivities'], num2_format)
		worksheet_3.write(offset + 2, k+1, 	individual_cashflow_statement_sub_list[k]['ProfitLossForStatementOfCashFlows'], num2_format)
		worksheet_3.write(offset + 3, k+1, 	individual_cashflow_statement_sub_list[k]['AdjustmentsForReconcileProfitLoss'], num2_format)
		worksheet_3.write(offset + 4, k+1, 	individual_cashflow_statement_sub_list[k]['AdjustmentsForDepreciationExpense'], num2_format)
		worksheet_3.write(offset + 5, k+1, 	individual_cashflow_statement_sub_list[k]['CashFlowsFromUsedInInvestingActivities'], num2_format)
		worksheet_3.write(offset + 6, k+1, 	individual_cashflow_statement_sub_list[k]['PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities'], num2_format)
		worksheet_3.write(offset + 7, k+1, 	individual_cashflow_statement_sub_list[k]['PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities'], num2_format)
		worksheet_3.write(offset + 8, k+1, 	individual_cashflow_statement_sub_list[k]['PurchaseOfInvestmentProperty'], num2_format)
		worksheet_3.write(offset + 9, k+1,  individual_cashflow_statement_sub_list[k]['ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities'], num2_format)
		worksheet_3.write(offset + 10, k+1, individual_cashflow_statement_sub_list[k]['ProceedsFromSalesOfIntangibleAssetsClassifiedAsInvestingActivities'], num2_format)
		worksheet_3.write(offset + 11, k+1,	individual_cashflow_statement_sub_list[k]['ProceedsFromSalesOfInvestmentProperty'], num2_format) 
		worksheet_3.write(offset + 12, k+1, individual_cashflow_statement_sub_list[k]['CashFlowsFromUsedInFinancingActivities'], num2_format)
		worksheet_3.write(offset + 13, k+1, individual_cashflow_statement_sub_list[k]['ProceedsFromShortTermBorrowings'], num2_format)	
		worksheet_3.write(offset + 14, k+1, individual_cashflow_statement_sub_list[k]['DividendsPaidClassifiedAsFinancingActivities'], num2_format)
		worksheet_3.write(offset + 15, k+1, individual_cashflow_statement_sub_list[k]['CashAndCashEquivalentsAtBeginningOfPeriodCf'], num2_format)
		worksheet_3.write(offset + 16, k+1, individual_cashflow_statement_sub_list[k]['CashAndCashEquivalentsAtEndOfPeriodCf'], num2_format)

	worksheet_4.write(offset + 0, 0, "연결 현금흐름표", filter_format2)
	worksheet_4.set_column('A:A', 30)
	worksheet_4.write(offset + 1, 0, "영업활동현금흐름")
	worksheet_4.write(offset + 2, 0, " 당기순이익")
	worksheet_4.write(offset + 3, 0, " 당기순이익조정을 위한 가감")
	worksheet_4.write(offset + 4, 0, " 감가상각비")
	worksheet_4.write(offset + 5, 0, "투자활동현금흐름")
	worksheet_4.write(offset + 6, 0, " 유형자산의 취득")
	worksheet_4.write(offset + 7, 0, " 무형자산의 취득")
	worksheet_4.write(offset + 8, 0, " 투자부동산의 취득")
	worksheet_4.write(offset + 9, 0, " 유형자산의 처분")
	worksheet_4.write(offset + 10, 0, " 무형자산의 처분")
	worksheet_4.write(offset + 11, 0, " 투자부동산의 처분")
	worksheet_4.write(offset + 12, 0, "재무활동현금흐름")
	worksheet_4.write(offset + 13, 0, " 단기차입금의 증가")
	worksheet_4.write(offset + 14, 0, " 배당금의 지급")
	worksheet_4.write(offset + 15, 0, "기초현금및현금성자산")
	worksheet_4.write(offset + 16, 0, "기말현금및현금성자산")
	          
	worksheet_4.write(offset + 0, 1,	"2017.4Q", filter_format)
	worksheet_4.write(offset + 0, 2,	"2017.3Q", filter_format)
	worksheet_4.write(offset + 0, 3, 	"2017.2Q", filter_format)
	worksheet_4.write(offset + 0, 4, 	"2017.1Q", filter_format)
	worksheet_4.write(offset + 0, 5, 	"2016.4Q", filter_format)
	worksheet_4.write(offset + 0, 6, 	"2016.3Q", filter_format)
	worksheet_4.write(offset + 0, 7, 	"2016.2Q", filter_format)
	worksheet_4.write(offset + 0, 8, 	"2016.1Q", filter_format)
	
	
	for k in range(len(individual_cashflow_statement_sub_list_q)):
	
		worksheet_4.write(offset + 1, k+1,	individual_cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInOperatingActivities'], num2_format)
		worksheet_4.write(offset + 2, k+1, 	individual_cashflow_statement_sub_list_q[k]['ProfitLossForStatementOfCashFlows'], num2_format)
		worksheet_4.write(offset + 3, k+1, 	individual_cashflow_statement_sub_list_q[k]['AdjustmentsForReconcileProfitLoss'], num2_format)
		worksheet_4.write(offset + 4, k+1, 	individual_cashflow_statement_sub_list_q[k]['AdjustmentsForDepreciationExpense'], num2_format)
		worksheet_4.write(offset + 5, k+1, 	individual_cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInInvestingActivities'], num2_format)
		worksheet_4.write(offset + 6, k+1, 	individual_cashflow_statement_sub_list_q[k]['PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities'], num2_format)
		worksheet_4.write(offset + 7, k+1, 	individual_cashflow_statement_sub_list_q[k]['PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities'], num2_format)
		worksheet_4.write(offset + 8, k+1, 	individual_cashflow_statement_sub_list_q[k]['PurchaseOfInvestmentProperty'], num2_format)
		worksheet_4.write(offset + 9, k+1,  individual_cashflow_statement_sub_list_q[k]['ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities'], num2_format)
		worksheet_4.write(offset + 10, k+1, individual_cashflow_statement_sub_list_q[k]['ProceedsFromSalesOfIntangibleAssetsClassifiedAsInvestingActivities'], num2_format)
		worksheet_4.write(offset + 11, k+1,	individual_cashflow_statement_sub_list_q[k]['ProceedsFromSalesOfInvestmentProperty'], num2_format) 
		worksheet_4.write(offset + 12, k+1, individual_cashflow_statement_sub_list_q[k]['CashFlowsFromUsedInFinancingActivities'], num2_format)
		worksheet_4.write(offset + 13, k+1, individual_cashflow_statement_sub_list_q[k]['ProceedsFromShortTermBorrowings'], num2_format)	
		worksheet_4.write(offset + 14, k+1, individual_cashflow_statement_sub_list_q[k]['DividendsPaidClassifiedAsFinancingActivities'], num2_format)
		worksheet_4.write(offset + 15, k+1, individual_cashflow_statement_sub_list_q[k]['CashAndCashEquivalentsAtBeginningOfPeriodCf'], num2_format)
		worksheet_4.write(offset + 16, k+1, individual_cashflow_statement_sub_list_q[k]['CashAndCashEquivalentsAtEndOfPeriodCf'], num2_format)

	worksheet_raw = workbook.add_worksheet('지표_itooza')

	worksheet_raw.write(0, 0, "정보", filter_format)
	worksheet_raw.set_column('A:A', 15)
#	worksheet_raw.write(0, 1, "2017.12", filter_format)
#	worksheet_raw.write(0, 2, "2016.12", filter_format)
#	worksheet_raw.write(0, 3, "2015.12", filter_format)
#	worksheet_raw.write(0, 4, "2014.12", filter_format)
#	worksheet_raw.write(0, 5, "2013.12", filter_format)
#	worksheet_raw.write(0, 6, "2012.12", filter_format)
#	worksheet_raw.write(0, 7,  "2011.12", filter_format)
#	worksheet_raw.write(0, 8,  "2010.12", filter_format)
#	worksheet_raw.write(0, 9,  "2009.12", filter_format)
#	worksheet_raw.write(0, 10,  "2008.12", filter_format)
#	worksheet_raw.write(0, 11,  "2007.12", filter_format)
#	worksheet_raw.write(0, 12,  "2006.12", filter_format)

	worksheet_raw.write(1, 0, "주당순이익(연결)", filter_format2)
	worksheet_raw.write(2, 0, "주당순이익(개별)", filter_format2)
	worksheet_raw.write(3, 0, "PER", filter_format2)
	worksheet_raw.write(4, 0, "주당순자산", filter_format2)
	worksheet_raw.write(5, 0, "PBR", filter_format2)
	worksheet_raw.write(6, 0, "주당배당금", filter_format2)
	worksheet_raw.write(7, 0, "시가배당률(%)", filter_format2)
	worksheet_raw.write(8, 0, "ROE", filter_format2)
	worksheet_raw.write(9, 0, "순이익률", filter_format2)
	worksheet_raw.write(10, 0, "영업이익률", filter_format2)
	worksheet_raw.write(11, 0, "주가", filter_format2)

	worksheet_raw.write(14, 0, "분기별 정보", filter_format)
	worksheet_raw.write(15, 0, "주당순이익(연결)", filter_format2)
	worksheet_raw.write(16, 0, "주당순이익(개별)", filter_format2)
	worksheet_raw.write(17, 0, "PER", filter_format2)
	worksheet_raw.write(18, 0, "주당순자산", filter_format2)
	worksheet_raw.write(19, 0, "PBR", filter_format2)
	worksheet_raw.write(20, 0, "주당배당금", filter_format2)
	worksheet_raw.write(21, 0, "시가배당률(%)", filter_format2)
	worksheet_raw.write(22, 0, "ROE", filter_format2)
	worksheet_raw.write(23, 0, "순이익률", filter_format2)
	worksheet_raw.write(24, 0, "영업이익률", filter_format2)
	worksheet_raw.write(25, 0, "주가", filter_format2)

	#print(len(date_list))
	#print(len(date_list_q))
	for l in range(len(eps_connect_list)):
		worksheet_raw.write(0, l+1,	date_list[l+1], filter_format)
		worksheet_raw.write(1, l+1,	eps_connect_list[l], num2_format)
		worksheet_raw.write(2, l+1,	eps_individual_list[l], num2_format)
		worksheet_raw.write(3, l+1, per_list[l], num_format)
		worksheet_raw.write(4, l+1, bps_list[l], num2_format)
		worksheet_raw.write(5, l+1, pbr_list[l], num_format)
		worksheet_raw.write(6, l+1, dps_list[l], num2_format)
		worksheet_raw.write(7, l+1, dy_list[l],	 num_format)
		worksheet_raw.write(8, l+1, roe_list[l], num_format)
		worksheet_raw.write(9, l+1, net_margin_list[l],	 num_format)
		worksheet_raw.write(10, l+1, op_margin_list[l],	 num_format)
		worksheet_raw.write(11, l+1, stock_price_list[l],	 num2_format)

		worksheet_raw.write(14, l+1, date_list_q[l+1], filter_format)
		worksheet_raw.write(15, l+1, eps_connect_list_q[l], num2_format)
		worksheet_raw.write(16, l+1, eps_individual_list_q[l], num2_format)
		worksheet_raw.write(17, l+1, per_list_q[l], num_format)
		worksheet_raw.write(18, l+1, bps_list_q[l], num2_format)
		worksheet_raw.write(19, l+1, pbr_list_q[l], num_format)
		worksheet_raw.write(20, l+1, dps_list_q[l], num2_format)
		worksheet_raw.write(21, l+1, dy_list_q[l],	 num_format)
		worksheet_raw.write(22, l+1, roe_list_q[l], num_format)
		worksheet_raw.write(23, l+1, net_margin_list_q[l],	 num_format)
		worksheet_raw.write(24, l+1, op_margin_list_q[l],	 num_format)
		worksheet_raw.write(25, l+1, stock_price_list_q[l],	 num2_format)


class MyWindow(QMainWindow):
	def __init__(self):
		super().__init__()
		self.setupUI()
	
	def setupUI(self):
		self.setWindowTitle("Get financial info. from DART")
		self. setGeometry(800, 400, 350, 150)
		
		textLabel = QLabel("Message: ", self)
		textLabel.move(20, 100)	
		
		self.label = QLabel("", self)
		self.label.move(80, 100)
		self.label.resize(150, 30)
        
		# Label
		label2 = QLabel("검색기업", self)
		label2.move(20, 40)      
		
		# LineEdit
		self.lineEdit = QLineEdit("", self)
		self.lineEdit.move(100, 40)
		#self.lineEdit.textChanged.connect(self.lineEditChanged)
		
		btn1 = QPushButton("Download", self)
		btn1.move(220, 40)
		btn1.clicked.connect(self.btn1_clicked)
		
		btn2 = QPushButton("Run", self)
		btn2.move(220, 80)
		btn2.clicked.connect(self.btn2_clicked)
        
		btn3 = QPushButton("Quit", self)
		btn3.move(220, 120)
		btn3.clicked.connect(QApplication.quit)
        
		## StatusBar
		#self.statusBar = QStatusBar(self)
		#self.setStatusBar(self.statusBar)	

		menubar = self.menuBar()
		filemenu = menubar.addMenu('&File')
		
		Action = QAction('종료', self)
		Action.triggered.connect(QApplication.quit)
		filemenu.addAction(Action)
	
		Action = QAction('정보', self)
		Action.triggered.connect(self.info_clicked)
		filemenu.addAction(Action)
		

	def btn1_clicked(self):
		self.label.setText("Downloading...")
		QApplication.processEvents()
		#zip_test()
		download_files()
		self.label.setText("Done!!")

	def btn2_clicked(self):
		self.label.setText("Making excel file...")
		QApplication.processEvents()
		run_dart(self.lineEdit.text())
		self.label.setText("Done!!")
	
	def info_clicked(self):
		QMessageBox.about(self, "Version 0.0", "<a href='http://blog.naver.com/jaden-agent/221222659414'>blog.naver.com/jaden-agent</a>")
	
	#def lineEditChanged(self):
	#	self.statusBar.showMessage(self.lineEdit.text())


def qt_test():
	app = QApplication(sys.argv)
	#label = QLabel("Hello PyQt")
	#label.show()
	myWindow = MyWindow()
	myWindow.show()
	app.exec_()

def main():

	#qt_test()

	#download_files()
	#zip_test()

	corp = "삼성전자"
	#corp = "리노공업"

	app = QApplication(sys.argv)
	myWindow = MyWindow()
	myWindow.show()
	app.exec_()

# Main
if __name__ == "__main__":
	main()


