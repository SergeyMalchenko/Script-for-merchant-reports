import class_currency as basic
import openpyxl
import datetime


#---------------------------------------
# Excel writing



def whitebit_excel(i, df_turnover, df_turnover_refund, statement_sheet, check_sheet, date_from, date_to):
  eu_rate_fee = 0.015
  noneu_rate_fee = 0.035
  hold_rate = 0.1

  x = basic.Currency([i], df_turnover, df_turnover_refund)
  statement_sheet['K10'] = x.count_complete()  # quantity of complete transactions
  match i:
    case 'USD':
      # EU  USD turnover
      statement_sheet['C7'] = x.turnover_comp_EU()
      statement_sheet['C25'] = x.turnover_refund_EU()
      eu_fee = basic.fee(x.turnover_comp_EU(), eu_rate_fee)

      # NON EU turnover
      statement_sheet['S7'] = x.turnover_comp_NonEU()
      noneu_fee = basic.fee(x.turnover_comp_NonEU(), noneu_rate_fee)


      # hold
      hold_usd = basic.hold(x.turnover_comp_EU(), eu_fee, hold_rate) + basic.hold(x.turnover_comp_NonEU(), noneu_fee, hold_rate) # holback = 10% from turnover
      statement_sheet['C73'] = hold_usd

      # checking with Excel Formula
      check_sheet['E2'] = x.turnover_comp_EU()
      check_sheet['E3'] = x.turnover_comp_NonEU()
      check_sheet['E6'] = x.turnover_refund_EU() + x.turnover_refund_NonEU()
      check_sheet['F6'] = x.count_refund()

    case 'EUR':
      # EU EUR turnover
      statement_sheet['K7'] = x.turnover_comp_EU()
      statement_sheet['K25'] = x.turnover_refund_EU()
      statement_sheet['K22'] = x.count_refund()
      eu_fee = basic.fee(x.turnover_comp_EU(), eu_rate_fee) + x.count_complete()*0.25


      # NON EU turnover
      statement_sheet['W7'] = x.turnover_comp_NonEU()
      noneu_fee = basic.fee(x.turnover_comp_NonEU(), noneu_rate_fee)
      statement_sheet['W25'] = x.turnover_refund_NonEU()


      # hold
      hold_eur =  basic.hold(x.turnover_comp_EU(), eu_fee, hold_rate) + basic.hold(x.turnover_comp_NonEU(), noneu_fee, hold_rate)  # holback = 10% from turnover
      statement_sheet['C74'] = hold_eur
      # Quantity of transactions
      statement_sheet['K10'] = x.count_complete()

      # checking with Excel Formula
      check_sheet['E4'] = x.turnover_comp_EU()
      check_sheet['E5'] = x.turnover_comp_NonEU()
      check_sheet['E7'] = x.turnover_refund_EU() + x.turnover_refund_NonEU()
      check_sheet['F8'] = x.count_complete()
      check_sheet['F7'] = x.count_refund()

  # Checking with Excel formulas
  check_sheet['G2'] = '=SUM(SUMIFS(Data!$I:$I, Data!$H:$H,"USD", Data!$G:$G,"complete", Data!$M:$M, {"AUSTRIA","BELGIUM","BULGARIA","CROATIA","CYPRUS","CZECH REPUBLIC","DENMARK","ESTONIA","FINLAND","FRANCE","GERMANY","GREECE","HUNGARY","IRELAND","ITALY","LATVIA","LITHUANIA","LUXEMBOURG","MALTA","NETHERLANDS","POLAND","PORTUGAL","ROMANIA","SPAIN","SLOVAKIA","SLOVENIA","SWEDEN"}))'
  check_sheet['G3'] = '=SUM(SUMIFS(Data!$I:$I, Data!$H:$H,"USD", Data!$G:$G,"complete", Data!$M:$M, {"ALBANIA","ALGERIA","ANDORRA","ANGOLA","ANGUILLA","ANTIGUA AND BARBUDA","ARGENTINA","ARMENIA","ARUBA","AUSTRALIA","AZERBAIJAN","BAHAMAS","BAHRAIN","BANGLADESH","BARBADOS","BELARUS","BELIZE","BENIN","BERMUDA","BOSNIA AND HERZEGOVINA","BRUNEI DARUSSALAM","BURKINA FASO","BRAZIL","CAMBODIA","CAMEROON","CANADA","CAPE VERDE","CAYMAN ISLANDS","CHILE","COLOMBIA","COSTA RICA","COTE D\'IVOIRE","CURACAO","DJIBOUTI","DOMINICA","DOMINICAN REPUBLIC","ECUADOR","EGYPT","EL SALVADOR","GABON","GEORGIA","GHANA","GIBRALTAR","GRENADA","GUATEMALA","GUAM","GUYANA","GAMBIA","HAITI","HONDURAS","HONG KONG","ICELAND","INDIA","INDONESIA","ISRAEL","JAMAICA","JAPAN","JORDAN","KAZAKHSTAN","KENYA","KOREA, REPUBLIC OF","KOSOVO, REPUBLIC OF","KUWAIT","KYRGYZSTAN","LAO PEOPLE\'S DEMOCRATIC REPUBLIC","LESOTHO","LIECHTENSTEIN","MACAO","MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF","MADAGASCAR","MALAWI","MALAYSIA","MALDIVES","MAURITANIA","MAURITIUS","MEXICO","MOLDOVA, REPUBLIC OF","MONACO","MONGOLIA","MONTENEGRO","MONTSERRAT","MOROCCO","MOZAMBIQUE","NAMIBIA","NEW ZEALAND","NICARAGUA","NIGERIA","NORWAY","OMAN","PANAMA","PAPUA NEW GUINEA","PARAGUAY","PERU","PHILIPPINES","PUERTO RICO","QATAR","RWANDA","RUSSIAN FEDERATION","SAINT KITTS AND NEVIS","SAINT LUCIA","SAINT VINCENT AND THE GRENADINES","SAMOA","SAN MARINO","SAUDI ARABIA","SENEGAL","SERBIA","SEYCHELLES","SIERRA LEONE","SINGAPORE","SINT MAARTEN (DUTCH PART)","SOUTH AFRICA","SRI LANKA","SURINAME","SWAZILAND","SWITZERLAND","TAIWAN, PROVINCE OF CHINA","TAJIKISTAN","TANZANIA, UNITED REPUBLIC OF","THAILAND","TOGO","TRINIDAD AND TOBAGO","TUNISIA","TURKEY","TURKMENISTAN","TURKS AND CAICOS ISLANDS","UGANDA","UKRAINE","UNITED ARAB EMIRATES","UNITED KINGDOM","UNITED STATES","URUGUAY","UZBEKISTAN","VENEZUELA, BOLIVARIAN REPUBLIC OF","VIET NAM","VIRGIN ISLANDS, BRITISH","VIRGIN ISLANDS, U.S.","ZAMBIA","NIGER","BOTSWANA","GUINEA"}))'
  check_sheet['G4'] = '=SUM(SUMIFS(Data!$I:$I, Data!$H:$H,"EUR", Data!$G:$G,"complete", Data!$M:$M, {"AUSTRIA","BELGIUM","BULGARIA","CROATIA","CYPRUS","CZECH REPUBLIC","DENMARK","ESTONIA","FINLAND","FRANCE","GERMANY","GREECE","HUNGARY","IRELAND","ITALY","LATVIA","LITHUANIA","LUXEMBOURG","MALTA","NETHERLANDS","POLAND","ROMANIA","PORTUGAL","SPAIN","SLOVAKIA","SLOVENIA","SWEDEN"}))'
  check_sheet['G5'] = '=SUM(SUMIFS(Data!$I:$I, Data!$H:$H,"EUR", Data!$G:$G,"complete", Data!$M:$M, {"ALBANIA","ALGERIA","ANDORRA","ANGOLA","ANGUILLA","ANTIGUA AND BARBUDA","ARGENTINA","ARMENIA","ARUBA","AUSTRALIA","AZERBAIJAN","BAHAMAS","BAHRAIN","BANGLADESH","BARBADOS","BELARUS","BELIZE","BENIN","BERMUDA","BOSNIA AND HERZEGOVINA","BRUNEI DARUSSALAM","BURKINA FASO","BRAZIL","CAMBODIA","CAMEROON","CANADA","CAPE VERDE","CAYMAN ISLANDS","CHILE","COLOMBIA","COSTA RICA","COTE D\'IVOIRE","CURACAO","DJIBOUTI","DOMINICA","DOMINICAN REPUBLIC","ECUADOR","EGYPT","EL SALVADOR","GABON","GEORGIA","GHANA","GIBRALTAR","GRENADA","GUATEMALA","GUAM","GUYANA","GAMBIA","HAITI","HONDURAS","HONG KONG","ICELAND","INDIA","INDONESIA","ISRAEL","JAMAICA","JAPAN","JORDAN","KAZAKHSTAN","KENYA","KOREA, REPUBLIC OF","KOSOVO, REPUBLIC OF","KUWAIT","KYRGYZSTAN","LAO PEOPLE\'S DEMOCRATIC REPUBLIC","LESOTHO","LIECHTENSTEIN","MACAO","MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF","MADAGASCAR","MALAWI","MALAYSIA","MALDIVES","MAURITANIA","MAURITIUS","MEXICO","MOLDOVA, REPUBLIC OF","MONACO","MONGOLIA","MONTENEGRO","MONTSERRAT","MOROCCO","MOZAMBIQUE","NAMIBIA","NEW ZEALAND","NICARAGUA","NIGERIA","NORWAY","OMAN","PANAMA","PAPUA NEW GUINEA","PARAGUAY","PERU","PHILIPPINES","PUERTO RICO","QATAR","RWANDA","RUSSIAN FEDERATION","SAINT KITTS AND NEVIS","SAINT LUCIA","SAINT VINCENT AND THE GRENADINES","SAMOA","SAN MARINO","SAUDI ARABIA","SENEGAL","SERBIA","SEYCHELLES","SIERRA LEONE","SINGAPORE","SINT MAARTEN (DUTCH PART)","SOUTH AFRICA","SRI LANKA","SURINAME","SWAZILAND","SWITZERLAND","TAIWAN, PROVINCE OF CHINA","TAJIKISTAN","TANZANIA, UNITED REPUBLIC OF","THAILAND","TOGO","TRINIDAD AND TOBAGO","TUNISIA","TURKEY","TURKMENISTAN","TURKS AND CAICOS ISLANDS","UGANDA","UKRAINE","UNITED ARAB EMIRATES","UNITED KINGDOM","UNITED STATES","URUGUAY","UZBEKISTAN","VENEZUELA, BOLIVARIAN REPUBLIC OF","VIET NAM","VIRGIN ISLANDS, BRITISH","VIRGIN ISLANDS, U.S.","ZAMBIA","NIGER","BOTSWANA","GUINEA"}))'
  check_sheet['G6'] = '=SUM(SUMIFS(Data!$I:$I,Data!$H:$H,"USD",Data!$G:$G,"refund"))'
  check_sheet['G7'] = '=SUM(SUMIFS(Data!$I:$I,Data!$H:$H,"EUR",Data!$G:$G,"refund"))'
  check_sheet['I6'] = '=COUNTIFS(Data!$G:$G,"refund",Data!$H:$H,"USD")'
  check_sheet['I7'] = '=COUNTIFS(Data!$G:$G,"refund",Data!$H:$H,"EUR")'
  check_sheet['I8'] = '=COUNTIFS(Data!$G:$G,"complete")'

#  Period
  statement_sheet['I3'] = date_from
  statement_sheet['Q3'] = date_to






