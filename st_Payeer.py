# Класс потомок для Пеер
import class_currency as basic


class clPayeer(basic.Currency):

  def count_refund (self):
    refund_count_df = self.df_turnover_refund[(self.df_turnover_refund['Валюта'].isin(self.curr_name)) & (~self.df_turnover_refund['Платежная система платежной карты'].isin(self.cup))]  # fore payeer & adv don't count CUP refund
    refund_count_curr = refund_count_df['Стоимость операции'].count()  # кол-во рефандов by currency
    return refund_count_curr
#  -----------------------------------------------------------------------------

def payeer_excel(i, df_turnover, df_turnover_refund, statement_sheet, check_sheet, date_from, date_to):
  x = clPayeer([i], df_turnover, df_turnover_refund)
  match i:
    case "USD":
      statement_sheet['C8'] = x.turnover_comp()
      statement_sheet['C24'] = x.turnover_refund()
      statement_sheet['C21'] = x.count_refund()


      # checking with Excel Formula
      check_sheet['D3'] = x.turnover_comp()
      check_sheet['D5'] = x.turnover_refund()
      check_sheet['E5'] = x.count_refund()
    case "EUR":
      statement_sheet['K8'] = x.turnover_comp()
      statement_sheet['K24'] = x.turnover_refund()
      statement_sheet['K21'] = x.count_refund()

      check_sheet['D2'] = x.turnover_comp()
      check_sheet['D4'] = x.turnover_refund()
      check_sheet['E4'] = x.count_refund()

    case 'RUR':
      statement_sheet['S8'] = x.turnover_comp()
      statement_sheet['S24'] = x.turnover_refund()
      statement_sheet['S21'] = x.count_refund()

  # holdback return = 0


  statement_sheet['C58'] = '0'
  statement_sheet['K58'] = '0'

# Checking with Excel formulas

  check_sheet['F2'] = '=SUMIFS(Data!I:I,Data!G:G,Check!B2,Data!H:H,Check!C2)'
  check_sheet['F3'] = '=SUMIFS(Data!I:I,Data!G:G,Check!B3,Data!H:H,Check!C3)'
  check_sheet['F4'] = '=SUMIFS(Data!I:I,Data!G:G,Check!B4,Data!H:H,Check!C4)'
  check_sheet['F5'] = '=SUMIFS(Data!I:I,Data!G:G,Check!B5,Data!H:H,Check!C5)'
  check_sheet['H4'] = '=COUNTIFS(Data!G:G,Check!B4,Data!H:H,Check!C4,Data!P:P,"<>CHINA UNION PAY")'
  check_sheet['H5'] = '=COUNTIFS(Data!G:G,Check!B5,Data!H:H,Check!C5,Data!P:P,"<>CHINA UNION PAY")'
  #  Period
  statement_sheet['I3'] = date_from
  statement_sheet['Q3'] = date_to

