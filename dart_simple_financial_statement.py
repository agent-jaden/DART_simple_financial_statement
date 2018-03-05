#-*- coding:utf-8 -*-
# Read text files
import os
import xlsxwriter
import urllib.request
import zipfile
import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic

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


		#영업활동현금흐름
		if (word_list[2] == corp) and (word_list[10] == "ifrs_CashFlowsFromUsedInOperatingActivities"):
			if word_list[index].strip() != "":
				cashflow_sub_list['CashFlowsFromUsedInOperatingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#당기순이익(손실)
		elif (word_list[2] == corp) and (word_list[10] == "dart_ProfitLossForStatementOfCashFlows"):
			if word_list[index].strip() != "":
				cashflow_sub_list['ProfitLossForStatementOfCashFlows']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#당기순이익조정을 위한 가감
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_AdjustmentsForReconcileProfitLoss"):
			if word_list[index].strip() != "":
				cashflow_sub_list['AdjustmentsForReconcileProfitLoss']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#감가상각비
		elif (word_list[2] == corp) and (word_list[10] == "dart_AdjustmentsForDepreciationExpense"):
			if word_list[index].strip() != "":
				cashflow_sub_list['AdjustmentsForDepreciationExpense']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#투자활동현금흐름
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_CashFlowsFromUsedInInvestingActivities"):
			if word_list[index].strip() != "":
				cashflow_sub_list['CashFlowsFromUsedInInvestingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#유형자산의 취득
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities"):
			if word_list[index].strip() != "":
				cashflow_sub_list['PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#무형자산의 취득
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities"):
			if word_list[index].strip() != "":
				cashflow_sub_list['PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#투자부동산의 취득
		elif (word_list[2] == corp) and (word_list[10] == "dart_PurchaseOfInvestmentProperty"):
			if word_list[index].strip() != "":
				cashflow_sub_list['PurchaseOfInvestmentProperty']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#유형자산의 처분
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities"):
			if word_list[index].strip() != "":
				cashflow_sub_list['ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#무형자산의 처분
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_ProceedsFromSalesOfIntangibleAssetsClassifiedAsInvestingActivities"):
			if word_list[index].strip() != "":
				cashflow_sub_list['ProceedsFromSalesOfIntangibleAssetsClassifiedAsInvestingActivities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#투자부동산의 처분
		elif (word_list[2] == corp) and (word_list[10] == "dart_ProceedsFromSalesOfInvestmentProperty"):
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
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_DividendsPaidClassifiedAsFinancingActivities"):
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

def scrape_balance_sheet(raw_data, index, corp):

	balance_sheet_sub_list = {}
	
	balance_sheet_sub_list['CurrentAssets']						= -1.0
	balance_sheet_sub_list['CashAndCashEquivalents']			= -1.0
	balance_sheet_sub_list['ShortTermDeposits']					= -1.0
	balance_sheet_sub_list['TradeAndOtherCurrentReceivables']	= -1.0
	balance_sheet_sub_list['Inventories']						= -1.0
	balance_sheet_sub_list['NoncurrentAssets']					= -1.0
	balance_sheet_sub_list['LongTermDeposits']					= -1.0	
	balance_sheet_sub_list['PropertyPlantAndEquipment']			= -1.0
	balance_sheet_sub_list['InvestmentProperty']				= -1.0
	balance_sheet_sub_list['DeferredTaxAssets']					= -1.0
	balance_sheet_sub_list['Assets']							= -1.0
	balance_sheet_sub_list['CurrentLiabilities']				= -1.0
	balance_sheet_sub_list['TradeAndOtherCurrentPayables']		= -1.0
	balance_sheet_sub_list['ShortTermBorrowings']				= -1.0
	balance_sheet_sub_list['NoncurrentLiabilities']				= -1.0
	balance_sheet_sub_list['BondsIssued']						= -1.0
	balance_sheet_sub_list['LongTermBorrowingsGross']			= -1.0
	balance_sheet_sub_list['DeferredTaxLiabilities']			= -1.0
	balance_sheet_sub_list['Liabilities']						= -1.0
	balance_sheet_sub_list['IssuedCapital']						= -1.0
	balance_sheet_sub_list['SharePremium']						= -1.0
	balance_sheet_sub_list['RetainedEarnings']					= -1.0
	balance_sheet_sub_list['Equity']							= -1.0

	unit = 100000000.0

	for j in range(len(raw_data)):

		word_list = raw_data[j]

		#유동자산
		if (word_list[2] == corp) and (word_list[10] == "ifrs_CurrentAssets"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['CurrentAssets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#현금및현금성자산
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_CashAndCashEquivalents"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['CashAndCashEquivalents']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 단기금융상품
		elif (word_list[2] == corp) and ((word_list[10] == "dart_ShortTermDepositsNotClassifiedAsCashEquivalents") or (word_list[10] == "ifrs_OtherCurrentFinancialAssets")):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermDeposits']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#매출채권 및 기타유동채권
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_TradeAndOtherCurrentReceivables"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['TradeAndOtherCurrentReceivables']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#재고자산
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_Inventories"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['Inventories']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#비유동자산
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_NoncurrentAssets"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['NoncurrentAssets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 장기매도가능금융자산
		elif (word_list[2] == corp) and ((word_list[10] == "dart_LongTermDepositsNotClassifiedAsCashEquivalents") or (word_list[10] == "ifrs_OtherNoncurrentFinancialAssets")):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['LongTermDeposits']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#유형자산
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_PropertyPlantAndEquipment"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['PropertyPlantAndEquipment']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 무형자산
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_InvestmentProperty") or (word_list[10] == "dart_GoodwillGross")):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['InvestmentProperty']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 이연법인세자산
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_DeferredTaxAssets"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['DeferredTaxAssets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#자산총계
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_Assets"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['Assets']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#유동부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_CurrentLiabilities"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['CurrentLiabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#매입채무
		elif (word_list[2] == corp) and ((word_list[10] == "ifrs_TradeAndOtherCurrentPayables") or (word_list[10] == "dart_ShortTermTradePayables")):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['TradeAndOtherCurrentPayables']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#단기차입금
		elif (word_list[2] == corp) and ((word_list[10] == "dart_ShortTermBorrowings") or (word_list[10] == "ifrs_ShorttermBorrowings")):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['ShortTermBorrowings']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#비유동부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_NoncurrentLiabilities"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['NoncurrentLiabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 사채
		elif (word_list[2] == corp) and (word_list[10] == "dart_BondsIssued"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['BondsIssued']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#장기차입금
		elif (word_list[2] == corp) and (word_list[10] == "dart_LongTermBorrowingsGross"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['LongTermBorrowingsGross']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 이연법인세부채
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_DeferredTaxLiabilities"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['DeferredTaxLiabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#부채총계
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_Liabilities"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['Liabilities']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#자본금
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_IssuedCapital"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['IssuedCapital']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		# 주식발행초과금
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_SharePremium"):
			if word_list[index].strip() != "":
				balance_sheet_sub_list['SharePremium']	=	float(word_list[index].replace(',','').replace('\"',''))/unit
		#이익잉여금(결손금)
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_RetainedEarnings"):
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

		# 매출액 or 영업수익
		if (word_list[2] == corp) and (word_list[10] == "ifrs_Revenue"):
			if word_list[index].strip() != "":
				revenue = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 매출원가
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_CostOfSales"):
			if word_list[index].strip() != "":
				cost_of_sale = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 매출총이익
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_GrossProfit"):
			if word_list[index].strip() != "":
				gross_profit = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 판매비와 관리비
		elif (word_list[2] == corp) and (word_list[10] == "dart_TotalSellingGeneralAdministrativeExpenses"):
			if word_list[index].strip() != "":
				admin_expenses = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 영업이익
		elif (word_list[2] == corp) and (word_list[10] == "dart_OperatingIncomeLoss"):
			if word_list[index].strip() != "":
				op_income = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 법인세비용차감전순이익
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_ProfitLossBeforeTax"):
			if word_list[index].strip() != "":
				profit_before_tax = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 법인세비용
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_IncomeTaxExpenseContinuingOperations"):
			if word_list[index].strip() != "":
				income_tax = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 당기순이익
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_ProfitLoss"):
			if word_list[index].strip() != "":
				profit = float(word_list[index].replace(",","").replace('\"',''))/unit
		# 기본주당순이익
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_BasicEarningsLossPerShare"):
			if word_list[index].strip() != "":
				basic_eps = float(word_list[index].replace(",","").replace('\"',''))
		#기타수익
		elif (word_list[2] == corp) and (word_list[10] == "dart_OtherGains"):
			if word_list[index].strip() != "":
				other_gain = float(word_list[index].replace(",","").replace('\"',''))/unit
		#기타비용
		elif (word_list[2] == corp) and (word_list[10] == "dart_OtherLosses"):
			if word_list[index].strip() != "":
				other_loss = float(word_list[index].replace(",","").replace('\"',''))/unit
		#금융수익
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_FinanceIncome"):
			if word_list[index].strip() != "":
				finance_income = float(word_list[index].replace(",","").replace('\"',''))/unit
		#금융비용
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_FinanceCosts"):
			if word_list[index].strip() != "":
				finance_cost = float(word_list[index].replace(",","").replace('\"',''))/unit
		#지분법이익
		elif (word_list[2] == corp) and (word_list[10] == "ifrs_ShareOfProfitLossOfAssociatesAndJointVenturesAccountedForUsingEquityMethod"):
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

	#2017 3Q
	file_name = "2017.3Q/2017_3분기보고서_04_현금흐름표_연결_20180131.txt"
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

	#2017 3Q
	file_name = "2017.3Q/2017_3분기보고서_04_현금흐름표_20180131.txt"
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

	#2017 3Q
	file_name = "2017.3Q/2017_3분기보고서_01_재무상태표_연결_20180131.txt"
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

	#2017 3Q
	file_name = "2017.3Q/2017_3분기보고서_01_재무상태표_20180131.txt"
	raw_data = read_raw_data(file_name)
	balance_sheet_sub_list.append(scrape_balance_sheet(raw_data, 12, corp))

	return balance_sheet_sub_list
	
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

	#2017 3Q
	if mode == 0:
		file_name = "2017.3Q/2017_3분기보고서_02_손익계산서_연결_20180131.txt"
	else:
		file_name = "2017.3Q/2017_3분기보고서_03_포괄손익계산서_연결_20180131.txt"
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

	#2017 3Q
	if mode == 0:
		file_name = "2017.3Q/2017_3분기보고서_02_손익계산서_20180131.txt"
	else:
		file_name = "2017.3Q/2017_3분기보고서_03_포괄손익계산서_20180131.txt"
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
		income_statement_sub_list[3][i] = income_statement_sub_list[3][i] - income_statement_sub_list[2][i] - income_statement_sub_list[1][i] - income_statement_sub_list[0][i]

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
		income_statement_sub_list[3][i] = income_statement_sub_list[3][i] - income_statement_sub_list[2][i] - income_statement_sub_list[1][i] - income_statement_sub_list[0][i]

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
	individual_balance_sheet_sub_list		= get_individual_balance_sheet_year(corp)
	individual_balance_sheet_sub_list_q		= get_individual_balance_sheet_quarter(corp)

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
	worksheet_1.write(0, 1,		"2017.3Q누적", filter_format)
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
	worksheet_2.write(0, 1,		"2017.3Q", filter_format)
	worksheet_2.write(0, 2, 	"2017.2Q", filter_format)
	worksheet_2.write(0, 3, 	"2017.1Q", filter_format)
	worksheet_2.write(0, 4, 	"2016.4Q", filter_format)
	worksheet_2.write(0, 5, 	"2016.3Q", filter_format)
	worksheet_2.write(0, 6, 	"2016.2Q", filter_format)
	worksheet_2.write(0, 7, 	"2016.1Q", filter_format)
	
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
	
	worksheet_1.write(offset + 0, 1,	"2017.3Q누적", filter_format)
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
	
	worksheet_2.write(offset + 0, 1,	"2017.3Q", filter_format)
	worksheet_2.write(offset + 0, 2, 	"2017.2Q", filter_format)
	worksheet_2.write(offset + 0, 3, 	"2017.1Q", filter_format)
	worksheet_2.write(offset + 0, 4, 	"2016.4Q", filter_format)
	worksheet_2.write(offset + 0, 5, 	"2016.3Q", filter_format)
	worksheet_2.write(offset + 0, 6, 	"2016.2Q", filter_format)
	worksheet_2.write(offset + 0, 7, 	"2016.1Q", filter_format)
	
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
	worksheet_3.write(0, 1,	"2017.3Q누적", filter_format)
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
	worksheet_4.write(0, 1,		"2017.3Q", filter_format)
	worksheet_4.write(0, 2, 	"2017.2Q", filter_format)
	worksheet_4.write(0, 3, 	"2017.1Q", filter_format)
	worksheet_4.write(0, 4, 	"2016.4Q", filter_format)
	worksheet_4.write(0, 5, 	"2016.3Q", filter_format)
	worksheet_4.write(0, 6, 	"2016.2Q", filter_format)
	worksheet_4.write(0, 7, 	"2016.1Q", filter_format)
	
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
	
	worksheet_3.write(offset + 0, 1,	"2017.3Q누적", filter_format)
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
	
	worksheet_4.write(offset + 0, 1,	"2017.3Q", filter_format)
	worksheet_4.write(offset + 0, 2, 	"2017.2Q", filter_format)
	worksheet_4.write(offset + 0, 3, 	"2017.1Q", filter_format)
	worksheet_4.write(offset + 0, 4, 	"2016.4Q", filter_format)
	worksheet_4.write(offset + 0, 5, 	"2016.3Q", filter_format)
	worksheet_4.write(offset + 0, 6, 	"2016.2Q", filter_format)
	worksheet_4.write(offset + 0, 7, 	"2016.1Q", filter_format)
	
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
	worksheet_1.write(offset + 4, 0, "  매출채권")
	worksheet_1.write(offset + 5, 0, "  재고자산")
	worksheet_1.write(offset + 6, 0, "비유동자산")
	worksheet_1.write(offset + 7, 0, "  장기매도가능금융자산")
	worksheet_1.write(offset + 8, 0, "  유형자산")
	worksheet_1.write(offset + 9, 0, "  무형자산")
	worksheet_1.write(offset + 10, 0, "  이연법인세자산")
	worksheet_1.write(offset + 11, 0, "자산총계")
	worksheet_1.write(offset + 12, 0, "유동부채")
	worksheet_1.write(offset + 13, 0, "  매입채무")
	worksheet_1.write(offset + 14, 0, "  단기차입금")
	worksheet_1.write(offset + 15, 0, "비유동부채")
	worksheet_1.write(offset + 16, 0, "  사채")
	worksheet_1.write(offset + 17, 0, "  장기차입금")
	worksheet_1.write(offset + 18, 0, "  이연법인세부채")
	worksheet_1.write(offset + 19, 0, "부채총계")
	worksheet_1.write(offset + 20, 0, "  자본금")
	worksheet_1.write(offset + 21, 0, "  주식발행초과금")
	worksheet_1.write(offset + 22, 0, "  이익잉여금")
	worksheet_1.write(offset + 23, 0, "자본총계")
	
	worksheet_1.write(offset + 0, 1,	"2017.3Q누적", filter_format)
	worksheet_1.write(offset + 0, 2, 	"2016년", filter_format)
	worksheet_1.write(offset + 0, 3, 	"2015년", filter_format)
	worksheet_1.write(offset + 0, 4, 	"2014년", filter_format)
	worksheet_1.write(offset + 0, 5,	"2013년", filter_format)
	
	for k in range(len(balance_sheet_sub_list)):
	
		worksheet_1.write(offset + 1, k+1, balance_sheet_sub_list[k]['CurrentAssets'], num2_format)						
		worksheet_1.write(offset + 2, k+1, balance_sheet_sub_list[k]['CashAndCashEquivalents'], num2_format)			
		worksheet_1.write(offset + 3, k+1, balance_sheet_sub_list[k]['ShortTermDeposits'], num2_format)					
		worksheet_1.write(offset + 4, k+1, balance_sheet_sub_list[k]['TradeAndOtherCurrentReceivables'], num2_format)	
		worksheet_1.write(offset + 5, k+1, balance_sheet_sub_list[k]['Inventories'], num2_format)						
		worksheet_1.write(offset + 6, k+1, balance_sheet_sub_list[k]['NoncurrentAssets'], num2_format)					
		worksheet_1.write(offset + 7, k+1, balance_sheet_sub_list[k]['LongTermDeposits'], num2_format)					
		worksheet_1.write(offset + 8, k+1, balance_sheet_sub_list[k]['PropertyPlantAndEquipment'], num2_format)			
		worksheet_1.write(offset + 9, k+1, balance_sheet_sub_list[k]['InvestmentProperty'], num2_format)				
		worksheet_1.write(offset + 10, k+1, balance_sheet_sub_list[k]['DeferredTaxAssets'], num2_format)	
		worksheet_1.write(offset + 11, k+1, balance_sheet_sub_list[k]['Assets'], num2_format)							
		worksheet_1.write(offset + 12, k+1, balance_sheet_sub_list[k]['CurrentLiabilities'], num2_format)				
		worksheet_1.write(offset + 13, k+1, balance_sheet_sub_list[k]['TradeAndOtherCurrentPayables'], num2_format)		
		worksheet_1.write(offset + 14, k+1, balance_sheet_sub_list[k]['ShortTermBorrowings'], num2_format)				
		worksheet_1.write(offset + 15, k+1, balance_sheet_sub_list[k]['NoncurrentLiabilities'], num2_format)			
		worksheet_1.write(offset + 16, k+1, balance_sheet_sub_list[k]['BondsIssued'], num2_format)						
		worksheet_1.write(offset + 17, k+1, balance_sheet_sub_list[k]['LongTermBorrowingsGross'], num2_format)			
		worksheet_1.write(offset + 18, k+1, balance_sheet_sub_list[k]['DeferredTaxLiabilities'], num2_format)			
		worksheet_1.write(offset + 19, k+1, balance_sheet_sub_list[k]['Liabilities'], num2_format)						
		worksheet_1.write(offset + 20, k+1, balance_sheet_sub_list[k]['IssuedCapital'], num2_format)				
		worksheet_1.write(offset + 21, k+1, balance_sheet_sub_list[k]['SharePremium'], num2_format)						
		worksheet_1.write(offset + 22, k+1, balance_sheet_sub_list[k]['RetainedEarnings'], num2_format)					
		worksheet_1.write(offset + 23, k+1, balance_sheet_sub_list[k]['Equity'], num2_format)							
	

	#worksheet_6 = workbook.add_worksheet('연결_재무상태표_분기')
	
	worksheet_2.write(offset + 0, 0, "연결 재무상태표", filter_format2)
	worksheet_2.set_column('A:A', 30)
	worksheet_2.write(offset + 1, 0, "유동자산")
	worksheet_2.write(offset + 2, 0, "  현금및현금성자산")
	worksheet_2.write(offset + 3, 0, "  단기금융상품")
	worksheet_2.write(offset + 4, 0, "  매출채권")
	worksheet_2.write(offset + 5, 0, "  재고자산")
	worksheet_2.write(offset + 6, 0, "비유동자산")
	worksheet_2.write(offset + 7, 0, "  장기매도가능금융자산")
	worksheet_2.write(offset + 8, 0, "  유형자산")
	worksheet_2.write(offset + 9, 0, "  무형자산")
	worksheet_2.write(offset + 10, 0, "  이연법인세자산")
	worksheet_2.write(offset + 11, 0, "자산총계")
	worksheet_2.write(offset + 12, 0, "유동부채")
	worksheet_2.write(offset + 13, 0, "  매입채무")
	worksheet_2.write(offset + 14, 0, "  단기차입금")
	worksheet_2.write(offset + 15, 0, "비유동부채")
	worksheet_2.write(offset + 16, 0, "  사채")
	worksheet_2.write(offset + 17, 0, "  장기차입금")
	worksheet_2.write(offset + 18, 0, "  이연법인세부채")
	worksheet_2.write(offset + 19, 0, "부채총계")
	worksheet_2.write(offset + 20, 0, "  자본금")
	worksheet_2.write(offset + 21, 0, "  주식발행초과금")
	worksheet_2.write(offset + 22, 0, "  이익잉여금")
	worksheet_2.write(offset + 23, 0, "자본총계")
	
	worksheet_2.write(offset + 0, 1,	"2017.3Q", filter_format)
	worksheet_2.write(offset + 0, 2, 	"2017.2Q", filter_format)
	worksheet_2.write(offset + 0, 3, 	"2017.1Q", filter_format)
	worksheet_2.write(offset + 0, 4, 	"2016.4Q", filter_format)
	worksheet_2.write(offset + 0, 5, 	"2016.3Q", filter_format)
	worksheet_2.write(offset + 0, 6, 	"2016.2Q", filter_format)
	worksheet_2.write(offset + 0, 7, 	"2016.1Q", filter_format)
	
	
	for k in range(len(balance_sheet_sub_list_q)):
	
		worksheet_2.write(offset + 1, k+1, balance_sheet_sub_list_q[k]['CurrentAssets'], num2_format)						
		worksheet_2.write(offset + 2, k+1, balance_sheet_sub_list_q[k]['CashAndCashEquivalents'], num2_format)			
		worksheet_2.write(offset + 3, k+1, balance_sheet_sub_list_q[k]['ShortTermDeposits'], num2_format)					
		worksheet_2.write(offset + 4, k+1, balance_sheet_sub_list_q[k]['TradeAndOtherCurrentReceivables'], num2_format)	
		worksheet_2.write(offset + 5, k+1, balance_sheet_sub_list_q[k]['Inventories'], num2_format)						
		worksheet_2.write(offset + 6, k+1, balance_sheet_sub_list_q[k]['NoncurrentAssets'], num2_format)					
		worksheet_2.write(offset + 7, k+1, balance_sheet_sub_list_q[k]['LongTermDeposits'], num2_format)						
		worksheet_2.write(offset + 8, k+1, balance_sheet_sub_list_q[k]['PropertyPlantAndEquipment'], num2_format)			
		worksheet_2.write(offset + 9, k+1, balance_sheet_sub_list_q[k]['InvestmentProperty'], num2_format)				
		worksheet_2.write(offset + 10, k+1, balance_sheet_sub_list_q[k]['DeferredTaxAssets'], num2_format)					
		worksheet_2.write(offset + 11, k+1, balance_sheet_sub_list_q[k]['Assets'], num2_format)							
		worksheet_2.write(offset + 12, k+1, balance_sheet_sub_list_q[k]['CurrentLiabilities'], num2_format)				
		worksheet_2.write(offset + 13, k+1, balance_sheet_sub_list_q[k]['TradeAndOtherCurrentPayables'], num2_format)		
		worksheet_2.write(offset + 14, k+1, balance_sheet_sub_list_q[k]['ShortTermBorrowings'], num2_format)				
		worksheet_2.write(offset + 15, k+1, balance_sheet_sub_list_q[k]['NoncurrentLiabilities'], num2_format)				
		worksheet_2.write(offset + 16, k+1, balance_sheet_sub_list_q[k]['BondsIssued'], num2_format)						
		worksheet_2.write(offset + 17, k+1, balance_sheet_sub_list_q[k]['LongTermBorrowingsGross'], num2_format)			
		worksheet_2.write(offset + 18, k+1, balance_sheet_sub_list_q[k]['DeferredTaxLiabilities'], num2_format)			
		worksheet_2.write(offset + 19, k+1, balance_sheet_sub_list_q[k]['Liabilities'], num2_format)						
		worksheet_2.write(offset + 20, k+1, balance_sheet_sub_list_q[k]['IssuedCapital'], num2_format)						
		worksheet_2.write(offset + 21, k+1, balance_sheet_sub_list_q[k]['SharePremium'], num2_format)						
		worksheet_2.write(offset + 22, k+1, balance_sheet_sub_list_q[k]['RetainedEarnings'], num2_format)					
		worksheet_2.write(offset + 23, k+1, balance_sheet_sub_list_q[k]['Equity'], num2_format)							
	

	#worksheet_7 = workbook.add_worksheet('개별_재무상태표_year')
	
	worksheet_3.write(offset + 0, 0, "개별 재무상태표", filter_format2)
	worksheet_3.set_column('A:A', 30)
	worksheet_3.write(offset + 1, 0, "유동자산")
	worksheet_3.write(offset + 2, 0, "  현금및현금성자산")
	worksheet_3.write(offset + 3, 0, "  단기금융상품")
	worksheet_3.write(offset + 4, 0, "  매출채권")
	worksheet_3.write(offset + 5, 0, "  재고자산")
	worksheet_3.write(offset + 6, 0, "비유동자산")
	worksheet_3.write(offset + 7, 0, "  장기매도가능금융자산")
	worksheet_3.write(offset + 8, 0, "  유형자산")
	worksheet_3.write(offset + 9, 0, "  무형자산")
	worksheet_3.write(offset + 10, 0, "  이연법인세자산")
	worksheet_3.write(offset + 11, 0, "자산총계")
	worksheet_3.write(offset + 12, 0, "유동부채")
	worksheet_3.write(offset + 13, 0, "  매입채무")
	worksheet_3.write(offset + 14, 0, "  단기차입금")
	worksheet_3.write(offset + 15, 0, "비유동부채")
	worksheet_3.write(offset + 16, 0, "  사채")
	worksheet_3.write(offset + 17, 0, "  장기차입금")
	worksheet_3.write(offset + 18, 0, "  이연법인세부채")
	worksheet_3.write(offset + 19, 0, "부채총계")
	worksheet_3.write(offset + 20, 0, "  자본금")
	worksheet_3.write(offset + 21, 0, "  주식발행초과금")
	worksheet_3.write(offset + 22, 0, "  이익잉여금")
	worksheet_3.write(offset + 23, 0, "자본총계")
	
	worksheet_3.write(offset + 0, 1,	"2017.3Q누적", filter_format)
	worksheet_3.write(offset + 0, 2, 	"2016년", filter_format)
	worksheet_3.write(offset + 0, 3, 	"2015년", filter_format)
	worksheet_3.write(offset + 0, 4, 	"2014년", filter_format)
	worksheet_3.write(offset + 0, 5,	"2013년", filter_format)
	
	for k in range(len(individual_balance_sheet_sub_list)):
	
		worksheet_3.write(offset + 1, k+1, individual_balance_sheet_sub_list[k]['CurrentAssets'], num2_format)						
		worksheet_3.write(offset + 2, k+1, individual_balance_sheet_sub_list[k]['CashAndCashEquivalents'], num2_format)			
		worksheet_3.write(offset + 3, k+1, individual_balance_sheet_sub_list[k]['ShortTermDeposits'], num2_format)					
		worksheet_3.write(offset + 4, k+1, individual_balance_sheet_sub_list[k]['TradeAndOtherCurrentReceivables'], num2_format)	
		worksheet_3.write(offset + 5, k+1, individual_balance_sheet_sub_list[k]['Inventories'], num2_format)						
		worksheet_3.write(offset + 6, k+1, individual_balance_sheet_sub_list[k]['NoncurrentAssets'], num2_format)					
		worksheet_3.write(offset + 7, k+1, individual_balance_sheet_sub_list[k]['LongTermDeposits'], num2_format)						
		worksheet_3.write(offset + 8, k+1, individual_balance_sheet_sub_list[k]['PropertyPlantAndEquipment'], num2_format)			
		worksheet_3.write(offset + 9, k+1, individual_balance_sheet_sub_list[k]['InvestmentProperty'], num2_format)				
		worksheet_3.write(offset + 10, k+1, individual_balance_sheet_sub_list[k]['DeferredTaxAssets'], num2_format)					
		worksheet_3.write(offset + 11, k+1, individual_balance_sheet_sub_list[k]['Assets'], num2_format)							
		worksheet_3.write(offset + 12, k+1, individual_balance_sheet_sub_list[k]['CurrentLiabilities'], num2_format)				
		worksheet_3.write(offset + 13, k+1, individual_balance_sheet_sub_list[k]['TradeAndOtherCurrentPayables'], num2_format)		
		worksheet_3.write(offset + 14, k+1, individual_balance_sheet_sub_list[k]['ShortTermBorrowings'], num2_format)				
		worksheet_3.write(offset + 15, k+1, individual_balance_sheet_sub_list[k]['NoncurrentLiabilities'], num2_format)				
		worksheet_3.write(offset + 16, k+1, individual_balance_sheet_sub_list[k]['BondsIssued'], num2_format)						
		worksheet_3.write(offset + 17, k+1, individual_balance_sheet_sub_list[k]['LongTermBorrowingsGross'], num2_format)			
		worksheet_3.write(offset + 18, k+1, individual_balance_sheet_sub_list[k]['DeferredTaxLiabilities'], num2_format)			
		worksheet_3.write(offset + 19, k+1, individual_balance_sheet_sub_list[k]['Liabilities'], num2_format)						
		worksheet_3.write(offset + 20, k+1, individual_balance_sheet_sub_list[k]['IssuedCapital'], num2_format)						
		worksheet_3.write(offset + 21, k+1, individual_balance_sheet_sub_list[k]['SharePremium'], num2_format)						
		worksheet_3.write(offset + 22, k+1, individual_balance_sheet_sub_list[k]['RetainedEarnings'], num2_format)					
		worksheet_3.write(offset + 23, k+1, individual_balance_sheet_sub_list[k]['Equity'], num2_format)							
	
	#worksheet_8 = workbook.add_worksheet('개별_재무상태표_분기')
	
	worksheet_4.write(offset + 0, 0, "개별 재무상태표", filter_format2)
	worksheet_4.set_column('A:A', 30)
	worksheet_4.write(offset + 1, 0, "유동자산")
	worksheet_4.write(offset + 2, 0, "  현금및현금성자산")
	worksheet_4.write(offset + 3, 0, "  단기금융상품")
	worksheet_4.write(offset + 4, 0, "  매출채권")
	worksheet_4.write(offset + 5, 0, "  재고자산")
	worksheet_4.write(offset + 6, 0, "비유동자산")
	worksheet_4.write(offset + 7, 0, "  장기매도가능금융자산")
	worksheet_4.write(offset + 8, 0, "  유형자산")
	worksheet_4.write(offset + 9, 0, "  무형자산")
	worksheet_4.write(offset + 10, 0, "  이연법인세자산")
	worksheet_4.write(offset + 11, 0, "자산총계")
	worksheet_4.write(offset + 12, 0, "유동부채")
	worksheet_4.write(offset + 13, 0, "  매입채무")
	worksheet_4.write(offset + 14, 0, "  단기차입금")
	worksheet_4.write(offset + 15, 0, "비유동부채")
	worksheet_4.write(offset + 16, 0, "  사채")
	worksheet_4.write(offset + 17, 0, "  장기차입금")
	worksheet_4.write(offset + 18, 0, "  이연법인세부채")
	worksheet_4.write(offset + 19, 0, "부채총계")
	worksheet_4.write(offset + 20, 0, "  자본금")
	worksheet_4.write(offset + 21, 0, "  주식발행초과금")
	worksheet_4.write(offset + 22, 0, "  이익잉여금")
	worksheet_4.write(offset + 23, 0, "자본총계")
	
	worksheet_4.write(offset + 0, 1,	"2017.3Q", filter_format)
	worksheet_4.write(offset + 0, 2, 	"2017.2Q", filter_format)
	worksheet_4.write(offset + 0, 3, 	"2017.1Q", filter_format)
	worksheet_4.write(offset + 0, 4, 	"2016.4Q", filter_format)
	worksheet_4.write(offset + 0, 5, 	"2016.3Q", filter_format)
	worksheet_4.write(offset + 0, 6, 	"2016.2Q", filter_format)
	worksheet_4.write(offset + 0, 7, 	"2016.1Q", filter_format)
	
	
	for k in range(len(individual_balance_sheet_sub_list_q)):
	
		worksheet_4.write(offset + 1, k+1, individual_balance_sheet_sub_list_q[k]['CurrentAssets'], num2_format)						
		worksheet_4.write(offset + 2, k+1, individual_balance_sheet_sub_list_q[k]['CashAndCashEquivalents'], num2_format)			
		worksheet_4.write(offset + 3, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermDeposits'], num2_format)					
		worksheet_4.write(offset + 4, k+1, individual_balance_sheet_sub_list_q[k]['TradeAndOtherCurrentReceivables'], num2_format)	
		worksheet_4.write(offset + 5, k+1, individual_balance_sheet_sub_list_q[k]['Inventories'], num2_format)						
		worksheet_4.write(offset + 6, k+1, individual_balance_sheet_sub_list_q[k]['NoncurrentAssets'], num2_format)					
		worksheet_4.write(offset + 7, k+1, individual_balance_sheet_sub_list_q[k]['LongTermDeposits'], num2_format)						
		worksheet_4.write(offset + 8, k+1, individual_balance_sheet_sub_list_q[k]['PropertyPlantAndEquipment'], num2_format)			
		worksheet_4.write(offset + 9, k+1, individual_balance_sheet_sub_list_q[k]['InvestmentProperty'], num2_format)				
		worksheet_4.write(offset + 10, k+1, individual_balance_sheet_sub_list_q[k]['DeferredTaxAssets'], num2_format)					
		worksheet_4.write(offset + 11, k+1, individual_balance_sheet_sub_list_q[k]['Assets'], num2_format)							
		worksheet_4.write(offset + 12, k+1, individual_balance_sheet_sub_list_q[k]['CurrentLiabilities'], num2_format)				
		worksheet_4.write(offset + 13, k+1, individual_balance_sheet_sub_list_q[k]['TradeAndOtherCurrentPayables'], num2_format)		
		worksheet_4.write(offset + 14, k+1, individual_balance_sheet_sub_list_q[k]['ShortTermBorrowings'], num2_format)				
		worksheet_4.write(offset + 15, k+1, individual_balance_sheet_sub_list_q[k]['NoncurrentLiabilities'], num2_format)				
		worksheet_4.write(offset + 16, k+1, individual_balance_sheet_sub_list_q[k]['BondsIssued'], num2_format)						
		worksheet_4.write(offset + 17, k+1, individual_balance_sheet_sub_list_q[k]['LongTermBorrowingsGross'], num2_format)			
		worksheet_4.write(offset + 18, k+1, individual_balance_sheet_sub_list_q[k]['DeferredTaxLiabilities'], num2_format)			
		worksheet_4.write(offset + 19, k+1, individual_balance_sheet_sub_list_q[k]['Liabilities'], num2_format)						
		worksheet_4.write(offset + 20, k+1, individual_balance_sheet_sub_list_q[k]['IssuedCapital'], num2_format)						
		worksheet_4.write(offset + 21, k+1, individual_balance_sheet_sub_list_q[k]['SharePremium'], num2_format)						
		worksheet_4.write(offset + 22, k+1, individual_balance_sheet_sub_list_q[k]['RetainedEarnings'], num2_format)					
		worksheet_4.write(offset + 23, k+1, individual_balance_sheet_sub_list_q[k]['Equity'], num2_format)							

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
	
	worksheet_1.write(offset + 0, 1,	"2017.3Q누적", filter_format)
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
	          
	worksheet_2.write(offset + 0, 1,	"2017.3Q", filter_format)
	worksheet_2.write(offset + 0, 2, 	"2017.2Q", filter_format)
	worksheet_2.write(offset + 0, 3, 	"2017.1Q", filter_format)
	worksheet_2.write(offset + 0, 4, 	"2016.4Q", filter_format)
	worksheet_2.write(offset + 0, 5, 	"2016.3Q", filter_format)
	worksheet_2.write(offset + 0, 6, 	"2016.2Q", filter_format)
	worksheet_2.write(offset + 0, 7, 	"2016.1Q", filter_format)
	
	
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
	
	worksheet_3.write(offset + 0, 1,	"2017.3Q누적", filter_format)
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
	          
	worksheet_4.write(offset + 0, 1,		"2017.3Q", filter_format)
	worksheet_4.write(offset + 0, 2, 	"2017.2Q", filter_format)
	worksheet_4.write(offset + 0, 3, 	"2017.1Q", filter_format)
	worksheet_4.write(offset + 0, 4, 	"2016.4Q", filter_format)
	worksheet_4.write(offset + 0, 5, 	"2016.3Q", filter_format)
	worksheet_4.write(offset + 0, 6, 	"2016.2Q", filter_format)
	worksheet_4.write(offset + 0, 7, 	"2016.1Q", filter_format)
	
	
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
		QMessageBox.about(self, "Version 0.0", "<a href='http://blog.naver.com/jaden-agent'>blog.naver.com/jaden-agent</a>")
	
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


