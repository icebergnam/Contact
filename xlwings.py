
import xlwings as xw

class Xlstest:
    def NewWorkbook(self, name):
        self.name = name
        print( name )
        wb = xw.Workbook()
        # xw.range('a1').value = 'Foo 1'
        # xw.range('a1').value



def main():
    # kim = Contact('홍길동','000-111-2222','icebergnam@hotmail.com','Seoul')
    # kim.print_info()
    wb = Xlstest()
    wb.NewWorkbook('kim')


if __name__ == "__main__":
    main()

