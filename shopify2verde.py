import pandas as pd
import requests
import csv
import os
import sys,getopt
import time
import numpy as np
import re

def print_help():
    help_str=' python shopify2verde.py -v FNAME_VERDE -s FNAME_SHOPIFY [-o SKUOPTION] \n\
            -v FNAME_VERDE  (--verde=FNAME_VERDE) where FNAME_VERDE is the path to the VERDE template .csv file. Should just be this file: https://secure-wms.com/ViaSub.WMS/ImportFiles/Order%20Import%20Template.xlsx. The script will *TRY* to download this file but if the URL has changed, you will need to have a copy of it in your local area.\n\
            -s FNAME_SHOPIFY (--shopify==FNAME_SHOPIFY) where FNAME_SHOPIFY is the path to the SHOPIFY template .csv file\n\
            -o SKUOPTION (--only=SKUOPTION) can be "mats" "playpens" "balls" "accessories", filters out sku list from shopify and transforms them into a singular SKU.\n\n\
            For example, playpen-mat-balls-blue with "-only mats" option transforms it to "elite-play-mat-v2"\n\
            Another example, playpen-blue with "-only mats" option will just have the row dropped and ignored (won\'t be put into the verde output)'
    print(help_str)

def is_po_box(txt):
    po_box_detected = False
    if isinstance(txt,str):
        po_box_detected = re.search(r'(?:post(?:al)? (?:office )?|p[. ]?o\.? )?box',txt,re.IGNORECASE|re.MULTILINE)
    return True if po_box_detected else False

def process_args(opts, args):
    verde_fname = None
    shopify_fname = None
    only_sku_of = None
    for opt, arg in opts:
        if opt == "-h":
            print_help()
        elif opt in ("-v", "--verde"):
            verde_fname = arg
        elif opt in ("-s", "--shopify"):
            shopify_fname = arg
        elif opt in ("-o", "--only"):
            only_sku_of = arg
    return (verde_fname, shopify_fname, only_sku_of)


def download_verde_template(verde_xlsx_fname):
    if not os.path.isfile("./"+verde_xlsx_fname):
        verde_order_template = "https://secure-wms.com/ViaSub.WMS/ImportFiles/Order%20Import%20Template.xlsx"
        verde_template_content = requests.get(verde_order_template).content
        output = open(verde_xlsx_fname,'wb')
        output.write(verde_template_content)
        output.close()

def shop2verde(shop_pd,verde_pd, only_sku_of, outfile):
    shop_pd['Shipping Zip']=shop_pd['Shipping Zip'].astype(str).str.zfill(5)
    for idx, row in shop_pd.iterrows():
        if (row['Financial Status']!="paid"):
            # not paid yet / cancelled / refunded, so we're not gonna order anything for them.
            continue
        curr_row_dict = {}
        # convert all SKUs to the specified accessory
        # TODO - when i do a gui have a list of possible items to pick, like "elite-play-mat" "elite-play-mat-v2"
        curr_shop_sku = row['Lineitem sku']
        if only_sku_of=="mats":
            if "playpen-mat" in curr_shop_sku or "elite-play-mat" in curr_shop_sku:
                curr_row_dict['SKU'] = 'elite-play-mat-v2'
            else:
                continue #this order has no mat, so we don't need to fork the mat fulfillment to 3PL, skip this row.
        elif only_sku_of=="balls":
            if "playpen-mat-balls" in curr_shop_sku:
                curr_row_dict['SKU'] = 'pitballs-100' #only set of 100 balls comes with the playpen+mat+balls bundle
            elif "balls" in curr_shop_sku:
                curr_row_dict['SKU'] = curr_shop_sku #if they ordered balls by themselves.
            else:
                continue #not part of ppe+mat+balls bundle and customer didn't order balls by themselves, so there's no balls in this order to fork to 3PL.
        elif only_sku_of=="playpens":
            if "playpen-mat" in curr_shop_sku:
                curr_row_dict['SKU'] = "playpen-blue" if "blue" in curr_shop_sku else "playpen-red" #only 2 colors for now. #TODO - change this if >2 colors.
            elif "playpen" in curr_shop_sku:
                curr_row_dict['SKU'] = curr_shop_sku #not a bundle, so customer just ordered a playpen in which case we just use the SKU customer already has.
            else:
                continue #this is a non-playpen order and thus can be skipped.
        elif only_sku_of!="accessories":
            curr_row_dict['SKU'] = curr_shop_sku #no SKU filtering needed.


        curr_row_dict['ReferenceNumber'] = row['Name']
        curr_row_dict['ShipCarrier'] = "RateShop"
        curr_row_dict['ShipService'] = "RateShop w/SmartPost- RS01"
        curr_row_dict['ShipToAddress1'] = row['Shipping Address1']
        curr_row_dict['ShipToAddress2'] = row['Shipping Address2']
        curr_row_dict['ShipToEmail'] = row['Email']
        curr_row_dict['ShipTo Name'] = row['Shipping Name']

        curr_row_dict['ShipToCity'] = row['Shipping City']
        curr_row_dict['ShipToState'] = row['Shipping Province']
        curr_row_dict['ShipToZip'] = row['Shipping Zip']
        curr_row_dict['ShipToCountry'] = row['Shipping Country']
        phone = str(row['Shipping Phone'])
        phone = phone.replace("(","").replace(")","").replace(" ","").replace("-","")
        if len(phone) == 11:
            phone = phone[1:]
        curr_row_dict['ShipToPhone'] = phone
        curr_row_dict['Quantity'] = row['Lineitem quantity']

        # various warning checks for orders
        if (row['Risk Level']!="Low"):
            # TODO - just auto-drop medium or high risk orders??
            print("WARNING: Order " + row['Name'] + " has a \"" + row['Risk Level'] + "\" risk level! Please make sure this is something we want to import and not delete!")
        if (is_po_box(row['Shipping Address1']) or is_po_box(row['Shipping Address2'])):
            print("WARNING, ORDER " + row['Name'] + " HAS A PO BOX. WE CANNOT SHIP TO PO BOXES. PLEASE FIX THIS BEFORE IMPORTING IT TO VERDE!!!")

        if only_sku_of=="accessories": 
            if "playpen-mat" in curr_shop_sku or "elite-play-mat" in curr_shop_sku:
                curr_row_dict['SKU'] = 'elite-play-mat-v2' #there is at least a mat in the order
                verde_pd = verde_pd.append(curr_row_dict, ignore_index=True)
            if "playpen-mat-balls" in curr_shop_sku or "pitballs-100" in curr_shop_sku:
                curr_row_dict['SKU'] = 'pitballs-100' #there are balls in the order.
                verde_pd = verde_pd.append(curr_row_dict, ignore_index=True)
            continue # don't double push
        verde_pd = verde_pd.append(curr_row_dict, ignore_index=True)

    #replace dataframe's NANs with empty string, as empty cells are NaNs
    verde_pd=verde_pd.replace(np.nan, '', regex=True)

    verde_pd.to_csv(outfile,index=False,header=None,sep='\t')

if __name__ == "__main__":
    # take in commandline arguments
    argv = sys.argv[1:]
    try:
        opts, args = getopt.getopt(argv,"hv:s:o:",["--verde=","--shopify=","--only="])
    except getopt.GetoptError:
        print_help()
        sys.exit(2)

    #run through each arg from commandline
    verde_fname, shopify_fname, only_sku_of= process_args(opts, args)

    # generate pd from verde and shopify data
    if (shopify_fname is None):
        print_help()
        print("--shopify= option is missing")
        sys.exit(2)
    if (verde_fname is None):
        #download verde .xlsx if no -v argument
        verde_xlsx_fname = "verde_template.xlsx"
        download_verde_template(verde_xlsx_fname)
        verde_pd = pd.read_excel(open(verde_xlsx_fname, 'rb'), header=1)  
    else:
        # just use the -v argument for the template file.
        verde_pd = pd.read_excel(open(verde_fname,'rb'), header=1)
    shop_pd = pd.read_csv(shopify_fname)
        
    # convert shopify to verde
    timestamp = time.strftime('%m-%d_%H%M%S', time.localtime())
    #TODO add a cmdline option for output file
    shop2verde(shop_pd=shop_pd,verde_pd=verde_pd, only_sku_of=only_sku_of, outfile="verde_"+timestamp+".txt")
    #TODO refactor this to be JD and VERDE orthogonally as options.
