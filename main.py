from intuitlib.client import AuthClient
from intuitlib.enums import Scopes
from quickbooks import QuickBooks
from quickbooks.objects.customer import Customer
from quickbooks.objects.invoice import Invoice
from quickbooks.objects.department import Department
from quickbooks.objects.item import Item
from quickbooks.objects.term import Term
from quickbooks.objects.classitem import ClassItem

from openpyxl import Workbook, load_workbook

from settings import *

class PythonQuickBooks(object):
    """
    PythonQuickBooks

    Main class that acts as in interfqace to the
    quickbooks library
    """

    def __init__(self):
        """
        Init function

        Create a client for the API and load all the clients
        to speed up the retrival process.

        TODO now it loads only 1000 clients, extend to all
        """

        super(PythonQuickBooks, self).__init__()
        self.client = self._create_client()
        self.customers = Customer.where(
            "Active=True",
            order_by='DisplayName',
            max_results=1000,
            qb=self.client)
        self.trovati = 0
        self.quanti = len(self.customers)

    def _create_client(self):
        auth_client = AuthClient(
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET,
            environment=ENVIRONMENT,
            redirect_uri='https://developer.intuit.com/v2/OAuth2Playground/RedirectUrl',
        )

        # Refresh token endpoint
        auth_client.refresh(refresh_token=REFRESH_TOKEN)

        client = QuickBooks(
            auth_client=auth_client,
            refresh_token=REFRESH_TOKEN,
            company_id=COMPANY_ID,
            minorversion=54
        )

        return client

    def _get_location(self, sales_name):
        try:
            location = Department.filter(Name=sales_name, qb=self.client)[0]
            return location
        except:
            return None

    # def list_customers(self):

    #     for c in self.customers:
    #         if c.SalesTermRef is not None:
    #             t = Term.get(id=c.SalesTermRef.value, qb=self.client)
    #             print(t)

    #     def single_item(self):
    #         item = Item.where("Name LIKE 'TEST%'", qb=self.client)[0]

    #         print(f"""
    # {item.Name}
    # BU:{item.ClassRef}
    # CAT:{item.Category}
    # SKU:{item.Sku}""")

    #         item.Name = "TEST Prodotto Modificato 2"
    #         item.ClassRef.Id = "400000000000583914"
    #         item.Category = 15
    #         #item.Sku = "Contratti Manutenzione"
    #         item.save(qb=self.client)

    # def get_categories(self):
    #     categories = Item.where("Type='Category'", qb=self.client)

    #     for cat in categories:
    #         print(f"""[{cat.Id}] {cat.Name} ({cat.ParentRef})""")

    # def get_items(self):
    #     categories = Item.where("Type='Category'", qb=self.client)

    #     for cat in categories:
    #         print(f"""[{cat.Id} {cat.Name}""")

    # def get_customer_details(self):
    #     customer = Customer.where("AlternatePhone.FreeFormNumber='IT01101890109'", qb=self.client)

        # print(f"{customer.SalesTermRef.Value}")
        # print(f"{customer}")

    # def get_classes(self):
    #     classes = ClassItem.all(qb=self.client)
    #     for c in classes:
    #         print(c)

    # def fix_customers(self):
    #     customers = Customer.filter(Active=True, max_results=1000, qb=self.client)
    #     # customers = Customer.where("Active = True", max_results=1000, qb=self.client)
    #     num = 0
    #     for c in customers:
    #         if c.DisplayName.startswith("Marino"):
    #             c.DisplayName = c.CompanyName
    #             c.FullyQualifiedName = c.CompanyName
    #             c.Suffix = "MARINO"
    #             c.save(qb=self.client)
    #             num += 1

        # print(f"TOTALE: {num}")


    def set_location_in_invoices(self):
        """
        Change Location value in invoice to Sales
        contained in client's suffix field

        TODO: make it scriptable to run into a cron job
        """
        invoices = Invoice.filter(max_results=1000,
                                  order_by="DocNumber DESC",
                                  qb=self.client)

        for i in invoices:
            if i.DepartmentRef is None:
                c = Customer.get(i.CustomerRef.value, qb=self.client)
                if c.Suffix != "":
                    location = self._get_location(sales_name=c.Suffix)
                    if location is not None:
                        # print(f"INVOICE {i.DocNumber} ASSIGNED TO {c.Suffix} LOCATION ID {location.Id}")
                        i.DepartmentRef = location.to_ref()
                        i.save(qb=self.client)
                    else:
                        print(f"[ERROR] Location IS NONE FOR {c.Suffix} OF {c}")
                else:
                    print(f"[ERROR] c.Suffix IS NONE FOR {c}")

    def _load_excel_file(self, path):
        """Load Excel file"""
        wb = load_workbook(path)
        return wb.active

    def _create_excel_file(self, title="Undefined"):
        """Load Excel file"""
        wb = Workbook()
        ws1 = wb.active
        ws1.title = title
        return wb

    def _get_sales_term(self, id):
        """
        Get information about Sales Terms

        Name
        DayOfMonthDue
        DueDays
        """
        return Term.get(id, qb=self.client)

    def _get_customer_from_vat(self, vat_number):
        """Get customer from VAT ID"""
        customer = None
        for c in self.customers:
            if c.AlternatePhone is not None:
                if c.AlternatePhone.FreeFormNumber == f"{vat_number}":
                    # print(f"TROVATO QUALCOSA per {vat_number}")
                    customer = c
        return customer

    def _get_customer_terms(self, customer):
        """Get Customer terms of payment"""
        if customer.SalesTermRef.value is not None:
            return Term.get(customer.SalesTermRef.value, qb=self.client)
        return None

    def _format_description(self, data, expenses=False):
        """Format description concatenating columns"""
        if expenses:
            return f"Rimborso spese incasso"
        return f"{data[7]} Modello {data[11]} Matricola {data[10]} CTR {data[5]}/{data[4]}"

    def _get_due_date(self, terms, ref):
        # print(f"{term.Name}\t{term.DayOfMonthDue}\t{term.DueNextMonthDays}\t{term.DueDays}")
        if terms.DueNextMonthDays is not None:
            months = terms.DueNextMonthDays / 30
            return f"=EOMONTH({ref}, {months})"
        return f"={ref}+{terms.DueDays}"

    def update_invoices_terms_and_due_date(self):
        """Read and update an Excel file to be used to import into QB"""
        path = '/Users/ale/Documents/git/python-quickbooks/data/invoices.xlsx'
        ws = self._load_excel_file(path)

        for row in ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
            if row[0] is not None:
                t = self._get_customer_terms(vat_number=row[0])


    # def list_terms(self):
    #     terms = Term.filter(Active=True, qb=self.client)
    #     for term in terms:
    #         print(f"{term.Name}\t{term.DayOfMonthDue}\t{term.DueNextMonthDays}\t{term.DueDays}")

    # def customer_check(self):
    #     for c in self.customers:
    #         if c.AlternatePhone is None:
    #             print(f"[{c.DisplayName}] PARTITA IVA VUOTA {c.PrimaryTaxIdentifier}")
    #         if c.PrimaryTaxIdentifier.startswith("XXX"):
    #             print(f"[{c.DisplayName}] PARTITA IVA ERRATA {c.PrimaryTaxIdentifier}")

    def import_invoices(self, input, output):
        """
        Import invoices into Quickbooks

        1. Open input file (INVOICES.xlsx) and set active sheet
        2. Create output file (TO_BE_IMPORTED.xlsx)
        2. For each row
            2.1 Select customer from VAT number (or CF if empty)
            2.2 Get customer Terms of Payment (ToP)
            2.3 Evaluate formula for due date based on ToP
            2.3 Write row into TO_BE_IMPORTED.xlsx
        """
        # 1
        ws = self._load_excel_file(input)
        # 2
        wb_out = self._create_excel_file(title="Fatture da importare")
        ws_out = wb_out.active
        self.headings(ws_out)
        row_no = 2

        for row in ws.iter_rows(min_row=2, min_col=1, max_col=12, values_only=True):
            # print(row[0])
            if row[0] is not None:
                # 2.1
                vat_id = f"IT{row[1]}"
                if row[1] == '':
                    vat_id = f"CF{row[2]}"

                customer = self._get_customer_from_vat(vat_number=vat_id)

                if customer is None:
                    print(f">>>>> CLIENTE NON TROVATO VAT ID: {vat_id}")
                else:
                    # 2.2
                    terms = self._get_customer_terms(customer)
                    if terms is None:
                        print(f">>>>> TERMINI DI PAGAMENTO NON TROVATI CLIENTE {customer}")

                    self._output(ws=ws_out, data=row, terms=terms, row_no=row_no)
                    if row[3] == 'S':
                        # add row for expenses
                        row_no += 1
                        self._output(
                            ws=ws_out,
                            data=row,
                            terms=terms,
                            row_no=row_no,
                            expenses=True
                        )
            row_no += 1

        wb_out.save(output)

    def _output(self, ws, data, terms, row_no, expenses=False):
        # Customer
        ws.cell(column=1, row=row_no, value=data[0])
        # Partita Iva
        vat = f"IT{data[1]}"
        if data[1] == '':
            vat = f"CF{data[2]}"
        ws.cell(column=2, row=row_no, value=vat)
        # Invoice no
        ws.cell(column=3, row=row_no, value='')
        # Invoice date
        ws.cell(column=4, row=row_no, value='')
        # Due date
        ref = ws.cell(column=4, row=row_no).coordinate
        ws.cell(column=5, row=row_no, value=self._get_due_date(terms, ref))
        # Terms
        ws.cell(column=6, row=row_no, value=f"{terms}")
        # Item product/service
        ws.cell(column=7, row=row_no, value=data[6])
        # Item description
        ws.cell(column=8, row=row_no, value=self._format_description(data, expenses=expenses))
        # Amount
        ws.cell(column=9, row=row_no, value=data[10])
        ws.cell(column=10, row=row_no, value='')

        if expenses:
            ws.cell(column=7, row=row_no, value="SP")
            ws.cell(column=9, row=row_no, value="4,50")

    def headings(self, ws):
        """
        Customer
        Partita Iva
        Invoice no                  EMPTY
        Invoice date                EMPTY
        Due date                    =
        Terms
        Item product/service
        Item description            EVALUATED
        Amount
        Item Tax Code
        """
        row = 1
        ws.cell(column=1, row=row, value='Customer')
        ws.cell(column=2, row=row, value='Partita Iva')
        ws.cell(column=3, row=row, value='Invoice no')
        ws.cell(column=4, row=row, value='Invoice date')
        ws.cell(column=5, row=row, value='Due date')
        ws.cell(column=6, row=row, value='Terms')
        ws.cell(column=7, row=row, value='Item product/service')
        ws.cell(column=8, row=row, value='Item description')
        ws.cell(column=9, row=row, value='Amount')
        ws.cell(column=10, row=row, value='Item Tax Code')

    def run(self):
        # self.get_customer_details()
        # self.import_invoices(input='data/invoices_origin.xlsx', output='data/output.xlsx')
        # print(f"TROVATI: {self.trovati} SU: {self.quanti}")
        self.set_location_in_invoices()




def main():

    pyqb = PythonQuickBooks()
    pyqb.run()


if __name__ == '__main__':
    main()