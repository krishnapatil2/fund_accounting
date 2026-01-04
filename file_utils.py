import os
import json
import sys
import shutil

def get_app_directory():
    """Get the application directory - works for both development and compiled EXE"""
    if getattr(sys, 'frozen', False):
        # Running as compiled EXE
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.abspath(__file__))

def ensure_consolidated_data_file():
    """Ensure consolidated_data.json exists with default data and ASIO Sub Fund 4 configurations"""
    app_dir = get_app_directory()
    consolidated_path = os.path.join(app_dir, "consolidated_data.json")
    
    # ASIO Sub Fund 4 default configurations
    asio_sf4_configs = {
        "asio_sf4_ft": {
            "RecordType": "",
            "RecordAction": "InsertUpdate",
            "KeyValue": "",
            "KeyValue.KeyName": "",
            "UserTranId1": "",
            "Portfolio": "ASIO - SF 4",
            "LocationAccount": "",
            "Strategy": "Default",
            "Investment": "",
            "Broker": "",
            "EventDate": "",
            "SettleDate": "",
            "ActualSettleDate": "",
            "Quantity": "",
            "Price": "",
            "PriceDenomination": "CALC",
            "CounterInvestment": "INR",
            "NetInvestmentAmount": "CALC",
            "NetCounterAmount": "CALC",
            "tradeFX": "",
            "ContractFxRateNumerator": "",
            "ContractFxRateDenominator": "",
            "ContractFxRate": "",
            "NotionalAmount": "",
            "FundStructure": "ASIO-SF4",
            "SpotDate": "",
            "PriceDirectly": "",
            "CounterFXDenomination": "CALC",
            "CounterTDateFx": "",
            "AccruedInterest": "",
            "InvestmentAccruedInterest": "",
            "Comments": "",
            "TradeExpenses.ExpenseNumber": "",
            "TradeExpenses.ExpenseCode": "",
            "TradeExpenses.ExpenseAmt": "",
            "TradeExpenses.ExpenseNumber1": "",
            "TradeExpenses.ExpenseCode1": "",
            "TradeExpenses.ExpenseAmt1": "",
            "TradeExpenses.ExpenseNumber2": "",
            "TradeExpenses.ExpenseCode2": "",
            "TradeExpenses.ExpenseAmt2": "",
            "NonCapExpenses.NonCapNumber": "",
            "NonCapExpenses.NonCapExpenseCode": "",
            "NonCapExpenses.NonCapAmount": "",
            "NonCapExpenses.NonCapCurrency": "",
            "NonCapExpenses.LocationAccount": "",
            "NonCapExpenses.NonCapLiabilityCode": "",
            "NonCapExpenses.NonCapPaymentType": "",
            "NonCapExpenses.NonCapNumber1": "",
            "NonCapExpenses.NonCapExpenseCode1": "",
            "NonCapExpenses.NonCapAmount1": "",
            "NonCapExpenses.NonCapCurrency1": "",
            "NonCapExpenses.LocationAccount1": "",
            "NonCapExpenses.NonCapLiabilityCode1": "",
            "NonCapExpenses.NonCapPaymentType1": "",
            "NonCapExpenses.NonCapNumber2": "",
            "NonCapExpenses.NonCapExpenseCode2": "",
            "NonCapExpenses.NonCapAmount2": "",
            "NonCapExpenses.NonCapCurrency2": "",
            "NonCapExpenses.LocationAccount2": "",
            "NonCapExpenses.NonCapLiabilityCode2": "",
            "NonCapExpenses.NonCapPaymentType2": ""
        },
        "asio_sf4_trading_code_mapping": {
            "FT": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT",
            "FT1": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT1",
            "FT2": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT2",
            "FT3": "Asio_Sub Fund_4_OHM_FO_DBSBK0000289_FT3",
        },
        "asio_sub_fund4_read_config": {
            "read_from_row": 1,
            "read_from_column": "A"
        }
    }
    
    if not os.path.exists(consolidated_path):
        # Create default consolidated data
        default_data = {
            "fund_filename_map": {
                "DBSBK0000033": {
                    "Fund Names": {"Default": "DIF-Class 1 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000036": {
                    "Fund Names": {"Default": "DIF-Class 2 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000038": {
                    "Fund Names": {"Default": "DIF-Class 3 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000042": {
                    "Fund Names": {"Default": "DIF-Class 5 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000044": {
                    "Fund Names": {"Default": "DIF-Class 6 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000043": {
                    "Fund Names": {"Default": "DIF-Class 7 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000049": {
                    "Fund Names": {"Default": "DIF-Class 8 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000050": {
                    "Fund Names": {"Default": "DIF-Class 9 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000051": {
                    "Fund Names": {"Default": "DIF-Class 10 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000052": {
                    "Fund Names": {"Default": "DIF-Class 11 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000074": {
                    "Fund Names": {"Default": "DIF-Class 12 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000179": {
                    "Fund Names": {"Default": "DIF-Class 13 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000189": {
                    "Fund Names": {"Default": "DIF-Class 14 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000192": {
                    "Fund Names": {"Default": "DIF-Class 15 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000214": {
                    "Fund Names": {"Default": "DIF-Class 16 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000216": {
                    "Fund Names": {"Default": "DIF-Class 17 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000217": {
                    "Fund Names": {"Default": "DIF-Class 18_Moon"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000232": {
                    "Fund Names": {"Default": "DIF-Class 19 Holding"},
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000247": {
                    "Fund Names": {
                        "CDS": "DIF-Class 21 CDS Holding",
                        "Default": "DIF-Class 21 Holding"
                    },
                    "Password": "AAGCD0792B"
                },
                "DBSBK0000178": {
                    "Fund Names": {"Default": "DGF-Cell 8"},
                    "Password": "AAICD1968M"
                },
                "BNPP00000458": {
                    "Fund Names": {"Default": "DGF-Cell 9"},
                    "Password": "AAICD2891H"
                },
                "DGF-Cell 10": {
                    "Fund Names": {"Default": "DGF-Cell 10"},
                    "Password": "AAICD3412C"
                },
                "BNPP00000480": {
                    "Fund Names": {"Default": "DGF-Cell 11"},
                    "Password": "AAICD6359G"
                },
                "BNPP00000488": {
                    "Fund Names": {"Default": "DGF-Cell 13"},
                    "Password": "AAICD7821M"
                },
                "BNPP00000540": {
                    "Fund Names": {"Default": "DGF-Cell 16"},
                    "Password": "AAJCD5624K"
                },
                "BNPP00000535": {
                    "Fund Names": {"Default": "DGF-Cell 17"},
                    "Password": "AAJCD4991K"
                },
                "DBSBK0000229": {
                    "Fund Names": {"Default": "DGF-Cell 18"},
                    "Password": "AAJCD6205G"
                },
                "DBSBK0000228": {
                    "Fund Names": {"Default": "DGF-Cell 19"},
                    "Password": "AAJCD6049E"
                },
                "DBSBK0000285": {
                    "Fund Names": {
                        "CDS": "DGF-Cell 23 CDS Holding",
                        "Default": "DGF-Cell 23 Holding"
                    },
                    "Password": "AAKCD6244Q"
                },
                "DBSBK0000299": {
                    "Fund Names": {"Default": "DGF-Cell 24 Holding"},
                    "Password": "AAKCD7324B"
                },
                "DBSBK0000353": {
                    "Fund Names": {"Default": "DGF-Cell 28 Holding"},
                    "Password": "AALCD1140J"
                },
                "DBSBK0000354": {
                    "Fund Names": {"Default": "DGF-Cell 29 Holding"},
                    "Password": "AALCD1141K"
                },
                "DBSBK0000380": {
                    "Fund Names": {"Default": "DGF-Cell 32 Holding"},
                    "Password": ""
                },
                "DBSBK0000356": {
                    "Fund Names": {"Default": "GlobalQ_AIF-III"},
                    "Password": ""
                },
                "DGF-Cell 36": {
                    "Fund Names": {"Default": "DGF-Cell 36"},
                    "Password": ""
                },
                "DGF-Cell 38": {
                    "Fund Names": {"Default": "DGF-Cell 38"},
                    "Password": ""
                }
            },
            "asio_recon_portfolio_mapping": {
                "ASIO - SF 3": "THE ASIO FUND VCC - SUB-FUND 3",
                "ASIO - SF 8_Golden": "THE ASIO FUND VCC - EMKAY BHARAT FUND - BHARATS GOLDEN DECADE",
                "ASIO-SF1_": "THE ASIO FUND VCC - FORT PANGEA",
                "ASIO - SF 8_Capital Builder": "THE ASIO FUND VCC - EMKAY BHARAT FUND",
                "ASIO - SF 8_Envi": "THE ASIO FUND VCC - EMKAY BHARAT FUND - ENVI",
                "ASIO - SF 4": "THE ASIO FUND VCC - SUB FUND 4",
                "ASIO - SF 10": "THE ASIO FUND VCC-HATHI CAPITAL INDIA GROWTH FUND I",
                "ASIO - SF 6": "THE ASIO FUND VCC - FRAGRANT HARBOUR INDIA OPPORTUNITIES FUND",
                "ASIO - SF 2": "",
                "ASIO - SF 7": ""
            },
            "asio_recon_format_1_headers": {
                "CLN_CODE": "Cln Code",
                "CLN_NAME": "Cln Name",
                "INSTR_CODE": "Instr Code",
                "INSTR_ISIN": "Instr ISIN",
                "INSTR_NAME": "Instr Name",
                "DEMAT_PHYSICAL": "Demat / Physical",
                "SETTLED_POSITION": "Settled Position",
                "PENDING_PURCHASE": "Pending Purchase",
                "PENDING_SALE": "Pending Sale",
                "BLOCKED": "Blocked",
                "BLOCKED_RECEIVABLES": "Blocked Receivables",
                "BLOCKED_PAYABLES": "Blocked Payables",
                "PENDING_CA_ENTITLEMENTS": "Pending CA Entitlements",
                "MARKET_PRICE_DATE": "Market Price Date",
                "MARKET_PRICE": "Market Price",
                "SETTLED_BLOCKED_MARKET_PRICE": "(Settled + Blocked * Market Price)",
                "SALEABLE": "Saleable",
                "CONTRACTUAL": "Contractual"
            },
            "asio_recon_format_2_headers": {
                "VALUE_DATE_AS_AT": "Value date as at",
                "SERVICE_LOCATION": "Service location",
                "SECURITIES_ACCOUNT_NAME": "Securities account name",
                "SECURITIES_ACCOUNT_NUMBER": "Securities account number",
                "SECURITY_TYPE": "Security type",
                "SECURITY_NAME": "Security name",
                "ISIN": "ISIN",
                "PLACE_OF_SETTLEMENT": "Place of settlement",
                "SETTLED_BALANCE": "Settled balance",
                "ANTICIPATED_DEBITS": "Anticipated debits",
                "AVAILABLE_BALANCE": "Available balance",
                "ANTICIPATED_CREDITS": "Anticipated credits",
                "TRADED_BALANCE": "Traded balance",
                "SECURITY_CURRENCY": "Security currency",
                "SECURITY_PRICE": "Security price",
                "INDICATIVE_SETTLED_VALUE": "Indicative settled value",
                "INDICATIVE_TRADED_VALUE": "Indicative traded value",
                "INDICATIVE_SETTLED_VALUE_INR": "Indicative settled value in INR",
                "INDICATIVE_TRADED_VALUE_INR": "Indicative traded value in INR"
            },
            "asio_recon_bhavcopy_headers": {
                "TRADDT": "TradDt",
                "BIZDT": "BizDt",
                "SGMT": "Sgmt",
                "SRC": "Src",
                "FININSTRMTP": "FinInstrmTp",
                "FININSTRMID": "FinInstrmId",
                "ISIN": "ISIN",
                "TCKRSYMB": "TckrSymb",
                "SCTYSRS": "SctySrs",
                "XPRYDT": "XpryDt",
                "FININSTRMACTLXPRYDT": "FininstrmActlXpryDt",
                "STRKPRIC": "StrkPric",
                "OPTNTP": "OptnTp",
                "FININSTRMnm": "FinInstrmNm",
                "OPNPRIC": "OpnPric",
                "HGHPRIC": "HghPric",
                "LWPRIC": "LwPric",
                "CLSPRIC": "ClsPric",
                "LASTPRIC": "LastPric"
            },
            "asio_geneva_headers": {
                "PORTFOLIO": "Portfolio",
                "INVESTMENT": "Investment",
                "INVESTMENT_DESCRIPTION": "Investment Description",
                "TRADED_QUANTITY": "Traded Quantity",
                "SETTLED_QUANTITY": "Settled Quantity",
                "CURRENCY": "Currency",
                "UNIT_COST": "Unit Cost",
                "COST_LOCAL": "Cost Local",
                "COST_BOOK": "Cost Book",
                "MARKET_PRICE_LOCAL": "Market Price Local",
                "MARKET_VALUE_LOCAL": "Market Value Local"
            },
            "lotsize_data": {
                "ALUMINI": 1000,
                "ALUMINIUM": 5000,
                "CASTOR": 50,
                "COCUDAKL": 100,
                "COPPER": 2500, 
                "CRUDEOIL": 100,
                "CRUDEOILM": 10,
                "DHANIYA": 50,
                "GOLD": 100,
                "GOLDM": 10,
                "GUARGUM5": 50,
                "GUARSEED10": 50,
                "JEERAUNJHA": 30,
                "LEAD": 5000,
                "LEADMINI": 1000,
                "NATGASMINI": 250,
                "NATGAS": 1250,
                "SILVER": 30,
                "SILVERM": 5,
                "SILVERMIC": 1,
                "TMCFGRNZM": 50,
                "ZINC": 5000,
                "ZINCMINI": 1000,
                "JEERAMINI": 10,
                "KAPAS": 200,
                "MCXBULLDEX": 30,
                "WTICRUDE": 100,
                "NATURALGAS": 1250,
                "GOLDTEN": 1,
                "ELECDMBL": 50,
                "ELECMBL": 50
            },
            "trade_headers": {
                "CLIENT_CODE": "Client\nCode",
                "EXCHANGE": "Exchange",
                "MKT_TYPE": "Mkt\nType",
                "DATE": "Date",
                "TRADE_TIME": "TradeTime",
                "SETL_NO": "SetlNo",
                "SCRIP_CODE": "Scrip Code",
                "SCRIP_NAME": "Scrip Name",
                "ISIN": "ISIN",
                "BUY_QTY": "Buy Qty",
                "BUY_RATE": "Buy Rate",
                "BUY_NET_RATE": "Buy\nNet Rate",
                "BUY_VALUE": "Buy Value",
                "SELL_QTY": "Sell Qty",
                "BALANCE": "Balance",
                "SELL_RATE": "Sell Rate",
                "SELL_NET_RATE": "Sell\nNet Rate",
                "SELL_VALUE": "Sell Value",
                "BREAK_EVEN": "BreakEven",
                "MARKET_RATE": "Market Rate",
                "TERMINAL": "Terminal",
                "DELY": "Dely"
            },
            "aafspl_car_future": {
                "Exchange": "MCX",
                "Issuer": "",
                "Ticker": "",
                "Cusip": "",
                "Sedol": "",
                "Isin": "",
                "AltKey1": "",
                "AltKey2": "",
                "BloombergID": "",
                "BloombergTicker": "",
                "BloombergUniqueID": "",
                "BloombergMarketSector": "",
                "SettleDays": "",
                "PricingFactor": "",
                "CurrentPriceDayRange": 1,
                "AssetType": "FT",
                "InvestmentType": "CMFT",
                "PriceDenomination": "INR",
                "BifurcationCurrency": "INR",
                "PrincipalCurrency": "INR",
                "IncomeCurrency": "INR",
                "RiskCurrency": "INR",
                "IssueCountry": "IN",
                "WithholdingTaxType": "Standard",
                "QDIEligibilityFlag": "Not Eligible",
                "SharesOutstanding": "",
                "SubIndustry": "",
                "SubIndustry2": "",
                "QuantityPrecision": 3,
                "InvestmentCrossZero": "Use Accounting Parameters",
                "PriceByPreference": "Currency",
                "ForwardPriceInterpolateFlag": 0,
                "PricingPrecision": 3,
                "FirstMarkDate": "01-01-2022",
                "LastMarkDate": "",
                "PriceList": "",
                "AutoGenerateMarks": 1,
                "CashSettlement": 1
            },
            "option_security": {
                "Exchange": "MCX",
                "Issuer": "",
                "Ticker": "",
                "Cusip": "",
                "Sedol": "",
                "Isin": "",
                "Riccode": "",
                "AltKey1": "",
                "AltKey2": "",
                "BloombergID": "",
                "BloombergTicker": "",
                "BloombergUniqueID": "",
                "BloombergMarketSector": "",
                "SettleDays": 0,
                "PricingFactor": 1,
                "CurrentPriceDayRange": 0,
                "AssetType": "OP",
                "InvestmentType": "CMFTOP",
                "PriceDenomination": "",
                "BifurcationCurrency": "INR",
                "PrincipalCurrency": "INR",
                "IncomeCurrency": "INR",
                "RiskCurrency": "INR",
                "IssueCountry": "IN",
                "WithholdingTaxType": "Standard",
                "QDIEligibilityFlag": "Not Eligible",
                "SharesOutstanding": "",
                "Beta": "",
                "SubIndustry": "",
                "SubIndustry2": "",
                "NonDeliverableCurrencyFlag": "",
                "QuantityPrecision": "",
                "InvestmentCrossZero": "",
                "PriceByPreference": "",
                "ForwardPriceInterpolateFlag": "",
                "SecFeeSchedule": "",
                "SecEligibleFlag": "",
                "WashSalesEligible": "",
                "ExerciseStyle": "European",
                "PricingPrecision": 15,
                "CashSettledFlag": 1
            },
            "car_trade_loader": {
                "RecordAction": "InsertUpdate",
                "KeyValue.KeyName": "",
                "UserTranId1": "",
                "Portfolio": "AAFSPL_CAR",
                "LocationAccount": "AAFSPL_CAR_Ventura Securities Ltd",
                "Strategy": "Default",
                "Broker": "",
                "PriceDenomination": "CALC",
                "CounterInvestment": "INR",
                "NetInvestmentAmount": "CALC",
                "NetCounterAmount": "CALC",
                "TradeFX": "",
                "ContractFxRateNumerator": "",
                "ContractFxRateDenominator": "",
                "ContractFxRate": "",
                "NotionalAmount": "",
                "FundStructure": "CALC",
                "SpotDate": "",
                "PriceDirectly": "",
                "CounterFXDenomination": "CALC",
                "CounterTDateFx": "",
                "AccruedInterest": "",
                "InvestmentAccruedInterest": "",
                "TradeExpenses.ExpenseNumber": "",
                "TradeExpenses.ExpenseCode": "",
                "TradeExpenses.ExpenseAmt": "",
                "NonCapExpenses.NonCapNumber": "",
                "NonCapExpenses.NonCapExpenseCode": "",
                "NonCapExpenses.NonCapAmount": "",
                "NonCapExpenses.NonCapCurrency": "",
                "NonCapExpenses.LocationAccount": "",
                "NonCapExpenses.NonCapLiabilityCode": "",
                "NonCapExpenses.NonCapPaymentType": ""
            },
            "geneva_custodian_mapping": {
                "DIF-Class 5_Heer": "DOVETAIL INDIA FUND CLASS 5 SHARES",
                "DIF-Class 5_Moonrock_F&O": "DOVETAIL INDIA FUND CLASS 5 SHARES",
                "DIF-Class 3_Saankhya": "DOVETAIL INDIA FUND CLASS 3 Shares",
                "DIF-Class 11_Eagle": "DOVETAIL INDIA FUND CLASS 11 SHARES",
                "DGF-Cell 2": "DOVETAIL GLOBAL FUND PCC CELL 2",
                "DIF-Class 1": "DOVETAIL INDIA FUND",
                "DGF-Cell 23": "DOVETAIL GLOBAL FUND PCC - CELL 23",
                "DIF-Class 6_Amal Parikh": "DOVETAIL INDIA FUND CLASS 6 SHARES",
                "DIF-Class 21": "DOVETAIL INDIA FUND CLASS 21",
                "DGF-Cell 28": "DOVETAIL GLOBAL FUND PCC - CELL 28",
                "DGF-Cell 29": "DOVETAIL GLOBAL FUND PCC CELL 29",
                "DIF-Class 2_Agile": "DOVETAIL INDIA FUND CLASS 2 SHARES",
                "DIF-Class 19": "DOVETAIL INDIA FUND CLASS 19",
                "ASIO - SF 2": "THE ASIO FUND VCC - SUB-FUND 2",
                "ASIO - SF 3": "THE ASIO FUND VCC - SUB-FUND 3",
                "ASIO - SF 4": "THE ASIO FUND VCC - SUB-FUND 4",
                "ASIO - SF 7": "THE ASIO FUND VCC FORT CAPITAL ABSOLUTE ALTERNATES",
                "DIF-Class 18_Moon": "DOVETAIL INDIA FUND CLASS 18",
                "DIF-Class 12_F&O": "DOVETAIL INDIA FUND CLASS 12",
                "DIF-Class 12_Moon New F&O": "DOVETAIL INDIA FUND CLASS 12",
                "DIF-Class 6": "DOVETAIL INDIA FUND CLASS 6 SHARES",
                "DIF-Class 8_Discipline": "DOVETAIL INDIA FUND CLASS 8 SHARES",
                "GlobalQ_AIF-III": "GLOBAL Q FUND",
                "DGF-Cell 32": "DOVETAIL GLOBAL FUND PCC - CELL 32",
                "DIF-Class 8_Panda": "DOVETAIL INDIA FUND CLASS 8 SHARES",
                "DGF- Cell 36": "",
                "DGF- Cell 38": "",
                "DIF-Class 8_HBE Capital": "DOVETAIL INDIA FUND CLASS 8 SHARES"
            }
        }
        
        # Add ASIO Sub Fund 4 configurations to default data
        default_data.update(asio_sf4_configs)
        
        # Add filter configurations to default data
        default_data["mcx_group2_filters"] = ['Commodity Future Option', 'Commodity Option', 'Commodity Future']
        default_data["fno_group2_filters"] = ['Equity Option', 'Index Option', 'Index Future', 'Equity Future']
        
        # Save the default consolidated data file
        try:
            with open(consolidated_path, "w", encoding="utf-8") as f:
                json.dump(default_data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            # If there's an error creating the file, log but don't fail
            # The app should still start even if config file creation fails
            print(f"Warning: Could not create consolidated_data.json: {e}")
    else:
        # File exists - check and add ASIO Sub Fund 4 configurations if missing
        try:
            # Check if file is readable and not empty
            if os.path.getsize(consolidated_path) == 0:
                # File is empty, recreate with default data including ASIO SF4 configs
                default_data = {}
                default_data.update(asio_sf4_configs)
                # Add filter configurations
                default_data["mcx_group2_filters"] = ['Commodity Future Option', 'Commodity Option', 'Commodity Future']
                default_data["fno_group2_filters"] = ['Equity Option', 'Index Option', 'Index Future', 'Equity Future']
                with open(consolidated_path, "w", encoding="utf-8") as f:
                    json.dump(default_data, f, indent=4, ensure_ascii=False)
            else:
                # File exists and has content - read and update if needed
                with open(consolidated_path, "r", encoding="utf-8") as f:
                    consolidated_data = json.load(f)
                
                # Ensure consolidated_data is a dictionary
                if not isinstance(consolidated_data, dict):
                    # Invalid format, recreate with default data
                    consolidated_data = {}
                    consolidated_data.update(asio_sf4_configs)
                    # Add filter configurations
                    consolidated_data["mcx_group2_filters"] = ['Commodity Future Option', 'Commodity Option', 'Commodity Future']
                    consolidated_data["fno_group2_filters"] = ['Equity Option', 'Index Option', 'Index Future', 'Equity Future']
                    with open(consolidated_path, "w", encoding="utf-8") as f:
                        json.dump(consolidated_data, f, indent=4, ensure_ascii=False)
                else:
                    # Check if asio_sf4_ft exists, if not add it
                    updated = False
                    for key, config in asio_sf4_configs.items():
                        if key not in consolidated_data:
                            consolidated_data[key] = config
                            updated = True
                    
                    # Initialize filter configurations if they don't exist
                    default_mcx_filters = ['Commodity Future Option', 'Commodity Option', 'Commodity Future']
                    default_fno_filters = ['Equity Option', 'Index Option', 'Index Future', 'Equity Future']
                    
                    if "mcx_group2_filters" not in consolidated_data:
                        consolidated_data["mcx_group2_filters"] = default_mcx_filters
                        updated = True
                    
                    if "fno_group2_filters" not in consolidated_data:
                        consolidated_data["fno_group2_filters"] = default_fno_filters
                        updated = True
                    
                    # Save updated data if changes were made
                    if updated:
                        with open(consolidated_path, "w", encoding="utf-8") as f:
                            json.dump(consolidated_data, f, indent=4, ensure_ascii=False)
        except json.JSONDecodeError as e:
            # File exists but is corrupted/invalid JSON - backup and recreate
            try:
                # Try to backup corrupted file
                backup_path = consolidated_path + ".backup"
                if os.path.exists(consolidated_path):
                    shutil.copy2(consolidated_path, backup_path)
                
                # Recreate with default data including ASIO SF4 configs
                default_data = {}
                default_data.update(asio_sf4_configs)
                # Add filter configurations
                default_data["mcx_group2_filters"] = ['Commodity Future Option', 'Commodity Option', 'Commodity Future']
                default_data["fno_group2_filters"] = ['Equity Option', 'Index Option', 'Index Future', 'Equity Future']
                with open(consolidated_path, "w", encoding="utf-8") as f:
                    json.dump(default_data, f, indent=4, ensure_ascii=False)
                print(f"Warning: consolidated_data.json was corrupted. Recreated file. Backup saved as {backup_path}")
            except Exception as backup_error:
                print(f"Warning: Could not backup corrupted file: {backup_error}")
        except PermissionError as e:
            # Permission denied - log but don't fail
            print(f"Warning: Permission denied accessing consolidated_data.json: {e}")
        except Exception as e:
            # Any other error reading/updating - log but don't fail
            # The app should still start even if config update fails
            print(f"Warning: Could not update consolidated_data.json: {e}")
    
    return consolidated_path
