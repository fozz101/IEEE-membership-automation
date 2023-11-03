import fitz
import os
import time
from pypdf import PdfMerger
start_time = time.time()




invoices_directory = os.getcwd() + "\\invoices"
updated_invoices_directory = os.getcwd() + "\\updated_invoices"


try:
    os.mkdir(invoices_directory)
except OSError as error:
    pass
    #print(error)
try:
    os.mkdir(updated_invoices_directory)
except OSError as error:
    pass
    #print(error)


#print(os.listdir(invoices_directory))
for index, invoice in enumerate(os.listdir(invoices_directory)):
    print(index)
    print(invoice)
    doc = fitz.open(invoice)
    selected_pages_number = [number for number in range(len(doc) - 1)]
    doc.select(selected_pages_number)
    doc.save(updated_invoices_directory+"\\updated_"+invoice)

print("FINISHED !")
print("--- %s seconds ---" % (time.time() - start_time))


pdfs = os.listdir(updated_invoices_directory)

merger = PdfMerger()

for pdf in pdfs:
    merger.append(pdf)

merger.write("Merged_Invoices.pdf")
merger.close()