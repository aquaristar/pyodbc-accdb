import time
import pyodbc
#import pymssql

# Program description ===============================================================================================
 # @source  : qbasic.py
 # @desc    : This Program is for building Product XML files (abcsync.xml and hawsync.xml) from ERP and Web DB.
 #------------------------------------------------------------------------
 # VER  DATE         AUTHOR      DESCRIPTION
 # ---  -----------  ----------  -----------------------------------------
 # 1.0  2015.12.31   Michael 	 founder              	 
 # 1.1  2016.01.26   Michael 	 founder              	 
 # ----------- ----------  -----------------------------------------------
 # Python Standard Application Building
 # Copyright 2016 ABC Warehouse,  All rights reserved.

XML_TEMPLATE = """  <product_list>
						<product>
							<brand></brand>
							<name></name>
							<url></url>
							<sku></sku>
							<short_desc></short_desc>
							<long_desc></long_desc>
							<image></image>
							<category>/category>
							<Sub_category></Sub_category>
							<price></price>
							<saleprice></saleprice>
							<availability></availability>
							<mfn_Num></mfn_Num>
						</product>
					</product_list>  """
# Program description ===============================================================================================

# Function define ===================================================================================================
def writeOutputFile(content, file, append):
    if(append is True):
        f = open(file, 'ab')
    else:
        f = open(file, 'wb')
    f.write(content + b'\n')
    f.close()

def printForDebug(msg, isDebug):
	if isDebug :
		print msg
# Function define ===================================================================================================

# Global variable ===================================================================================================
FLG_DEBUG = False

FILENAME_ABC = "abcsync_ksi.xml"
FILENAME_HAW = "hawsync_ksi.xml"

URL_ABC		 = "http://www.abcwarehouse.com/product_catalog/pc_proddetails.asp~prod_id~"
URL_HAW		 = "http://www.hawthorneonlinestore.com/product_catalog/pc_proddetails.asp~prod_id~"

IMAGE_ABC 	 = "http://www.abcwarehouse.com/product_images/"
IMAGE_HAW 	 = "http://www.hawthorneonlinestore.com/product_images/"

xml_product			= ""
tag_product_abc     = ""
tag_product_haw     = ""
# Global variable ==================================================================================================



# Main Function Logic start ========================================================================================
start_time = time.time()
#connection = pyodbc.connect('DRIVER={SQL Server};SERVER=yourServer;DATABASE=yourDatabase;UID=yourUsername;PWD=yourPass')
#connection_sql = pyodbc.connect("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=D:\\JohnHartunian\\05.QuantelBasic2Python\\abc-content.accdb;")
#connection_erp = pyodbc.connect("DSN=TYLER_ISAM;UID=user;PWD=password")
#connection_erp = pyodbc.connect("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=D:\\JohnHartunian\\05.QuantelBasic2Python\\abc-content.accdb;")
connection_erp = pyodbc.connect("DSN=easysoft-bridge;UID=dbo;PWD=tuna123")    # running from pc
connection_sql = pyodbc.connect("DSN=content;UID=sa")                         #running from pc

cursor_main = connection_erp.cursor()
sql_main	=	"""	SELECT im.item_number, LEFT(im.dist_code,1) AS NOTHAWFLAG, im.product_type, im.description, im.department, im.uni_price_code, im.web_price, im.sell_price, im.list_price, im.status_code, im.dist_code, im.second_desc_size, bm.brand_desc, im.instock_flg
					FROM DA1_INVENTORY_MASTER AS im, da6_brand_master AS bm, da1_inv_fact_tag ift			
					WHERE im.web_enable LIKE 'Y'
					AND im.status_code NOT IN ('N', 'D', 'X', 'T')
					AND im.instock_flg IN ('Y','1','2','3','4','8','A','B')
					AND LEFT(im.dist_code,1) IN ('A','H', ' ')
					AND bm.brand = im.brand
					AND im.item_number = ift.item_number
					ORDER BY im.item_number"""
cursor_main.execute(sql_main)
results = cursor_main.fetchall()

#Start Loop in Results-----------------------------------------------------------------------------------------------
for row in results:
	#print (row.item_number)
	if (not row.BRAND_DESC) or (row.BRAND_DESC.strip() == ''):
		var_brand 		= ""
	else:
		var_brand 		= row.BRAND_DESC	
	var_name 			= var_brand + " " + row.DESCRIPTION	
	var_sku 			= row.ITEM_NUMBER

	tag_brand 			= "<brand>%s</brand>\n" % var_brand
	tag_name 			= "<name>%s</name>\n" % var_name	
	tag_sku 			= "<sku>%s</sku>\n" % var_sku
	
	var_price 			= ""
	var_saleprice 		= ""
	var_availability 	= "1"
	var_mfn_Num 		= ""
	var_category 		= ""
	var_Sub_category 	= ""

	# Description Part Start --------------------------------------------------------------------------------------		
	sql_description		= """ SELECT Top 1 title FROM ABCCatalog WHERE ItemNo LIKE '%s' """ % (row.ITEM_NUMBER)	
	cursor_description 	= connection_sql.cursor()
	cursor_description.execute(sql_description)
	result_description 	= cursor_description.fetchone()

	if result_description:
		var_short_desc 	= result_description.title
	else:
		var_short_desc 	= ""
	var_long_desc 		= ""

	tag_short_desc 		= "<short_desc>\n%s</short_desc>\n" % var_short_desc
	tag_long_desc 		= "<long_desc>\n%s</long_desc>\n" % var_long_desc
	# Description Part End   --------------------------------------------------------------------------------------

	# Category Part Start -----------------------------------------------------------------------------------------	
	sql_category = """	SELECT cm.cat_id, cm.cat_description, cm.parent_id, cm.product_type, cm.web_store
						FROM DA1_CATEGORY_MASTER cm
						WHERE cm.cat_id IN (
							SELECT parent_id FROM DA1_category_master WHERE product_type LIKE '%s'
						)
						OR  product_type LIKE '%s' """ % (row.PRODUCT_TYPE, row.PRODUCT_TYPE)
	
	cursor_category = connection_erp.cursor()
	cursor_category.execute(sql_category)
	results_category = cursor_category.fetchall()
	printForDebug ("\n"	, FLG_DEBUG)
	printForDebug ("\n"	, FLG_DEBUG)
	printForDebug ("*******************Start******************"	, FLG_DEBUG)
	printForDebug (results_category	, FLG_DEBUG)
	
	if not results_category:
		tag_category = ""
		tag_Sub_category = ""
	else:
		for row_category in results_category:
			var_Sub_category += row_category.CAT_DESCRIPTION + "|"
		var_Sub_category = var_Sub_category.rstrip('|')

		if (not results_category[0].PARENT_ID) or (results_category[0].PARENT_ID.strip() == ''):
			var_category = results_category[0].CAT_DESCRIPTION
		else:
			tmp_parent_id = results_category[0].PARENT_ID
			while True:	
				sql_root = "SELECT Top 1 cat_id, cat_description, parent_id, product_type FROM DA1_category_master WHERE cat_id LIKE '%s'" % tmp_parent_id
				cursor_root = connection_erp.cursor()
				cursor_root.execute(sql_root)
				result_root = cursor_root.fetchone()
				if (not result_root.PARENT_ID) or (result_root.PARENT_ID.strip == ''):
					var_category = result_root.CAT_DESCRIPTION
					break
				tmp_parent_id = result_root.PARENT_ID

	tag_category 	= "<category>%s</category>\n" % var_category
	tag_Sub_category= "<Sub_category>%s</Sub_category>\n" % var_Sub_category
	printForDebug (tag_category	, FLG_DEBUG)
	printForDebug (tag_Sub_category	, FLG_DEBUG)
	printForDebug ( "******************End*******************", FLG_DEBUG)
	# Category Part End ---------------------------------------------------------------------------------------

	# Price part Start ----------------------------------------------------------------------------------------
	#print row.UNI_PRICE_CODE, row.WEB_PRICE, row.LIST_PRICE, row.SELL_PRICE, row.DEPARTMENT
	if row.UNI_PRICE_CODE == 'L':
		if row.DEPARTMENT in ['A','D','L','M','O','OB','R','RB']:
			var_price = row.WEB_PRICE
			var_saleprice = row.WEB_PRICE
		else:
			var_price = row.LIST_PRICE
			var_saleprice = row.LIST_PRICE
	else:
		if row.WEB_PRICE != 0:
			var_price = row.WEB_PRICE
		else:
			var_price = row.SELL_PRICE

		if len(row.DIST_CODE) > 1 and row.DIST_CODE[1] == 'N':
			var_price = ''

		sql_prfile = "SELECT Top 1 alt_flag, promo_price FROM DA1_prfile WHERE key_branch LIKE 'WEB' AND key_item_number LIKE '%s'" % row.ITEM_NUMBER
		cursor_prfile = connection_erp.cursor()
		cursor_prfile.execute(sql_prfile)
		result_prfile = cursor_prfile.fetchone()

		if (result_prfile) and ((not result_prfile.ALT_FLAG) or (result_prfile.ALT_FLAG.strip == '')):
		 	var_saleprice = result_prfile.PROMO_PRICE
		else:
		 	var_saleprice = var_price

	tag_price 	  = "<price>%s</price>\n" % var_price
	tag_saleprice = "<saleprice>%s</saleprice>\n" % var_saleprice
	# Price part start ---------------------------------------------------------------------------------------

	# Avaiability part start ---------------------------------------------------------------------------------
	if row.INSTOCK_FLG in ['3','5','6','8']:
		var_availability = "0"

	if row.UNI_PRICE_CODE == 'A':
		var_availability = "0"

	if row.STATUS_CODE in ['A','AB']:
		var_availability = "0"

	if row.STATUS_CODE in ['N','D','X','T']:
		var_availability = "0"

	if row.DEPARTMENT in ['I','IE']:
		var_availability = "0"

	if (len(row.DIST_CODE) > 1) and (row.DIST_CODE[1] == 'N'):
	 	var_availability = "0"

	tag_availability = "<availability>%s</availability>\n" % var_availability
	# Avaiability part end ----------------------------------------------------------------------------------
	
	var_mfn_Num = row.DESCRIPTION.split(' ')[0]
	tag_mfn_Num = "<mfn_Num>%s</mfn_Num>\n" % var_mfn_Num

	if row.NOTHAWFLAG.strip() == '':
		printForDebug( "Both output", FLG_DEBUG)
		tag_url 			= "<url>%s</url>\n" % (URL_ABC + row.ITEM_NUMBER)
		tag_image 			= "<image>%s</image>\n" % (IMAGE_ABC + row.ITEM_NUMBER + "_detail.jpg")
		xml_product			= (tag_brand + tag_name + tag_url + tag_sku + tag_short_desc + tag_long_desc + tag_image + tag_category + tag_Sub_category + tag_price + tag_saleprice + tag_availability + tag_mfn_Num)
		tag_product_abc		+= ("<product>\n%s</product>\n" % xml_product)

		tag_url 			= "<url>%s</url>\n" % (URL_HAW + row.ITEM_NUMBER)
		tag_image 			= "<image>%s</image>\n" % (IMAGE_HAW + row.ITEM_NUMBER + "_detail.jpg")
		xml_product			= (tag_brand + tag_name + tag_url + tag_sku + tag_short_desc + tag_long_desc + tag_image + tag_category + tag_Sub_category + tag_price + tag_saleprice + tag_availability + tag_mfn_Num)
		tag_product_haw		+= ("<product>\n%s</product>\n" % xml_product)

	elif row.NOTHAWFLAG == 'A':
		printForDebug( "ABC output", FLG_DEBUG)
		tag_url 			= "<url>%s</url>\n" % (URL_ABC + row.ITEM_NUMBER)
		tag_image 			= "<image>%s</image>\n" % (IMAGE_ABC + row.ITEM_NUMBER + "_detail.jpg")
		xml_product			= (tag_brand + tag_name + tag_url + tag_sku + tag_short_desc + tag_long_desc + tag_image + tag_category + tag_Sub_category + tag_price + tag_saleprice + tag_availability + tag_mfn_Num)
		tag_product_abc		+= ("<product>\n%s</product>\n" % xml_product)

	elif row.NOTHAWFLAG == 'H':
		printForDebug( "Haw output", FLG_DEBUG)
		tag_url 			= "<url>%s</url>\n" % (URL_HAW + row.ITEM_NUMBER)
		tag_image 			= "<image>%s</image>\n" % (IMAGE_HAW + row.ITEM_NUMBER + "_detail.jpg")
		xml_product			= (tag_brand + tag_name + tag_url + tag_sku + tag_short_desc + tag_long_desc + tag_image + tag_category + tag_Sub_category + tag_price + tag_saleprice + tag_availability + tag_mfn_Num)
		tag_product_haw		+= ("<product>\n%s</product>\n" % xml_product)
#Start Loop in Results---------------------------------------------------------------------------

if tag_product_abc != "":
	tag_product_list = "<product_list>\n%s</product_list>" % tag_product_abc.encode('utf-8')
	writeOutputFile(tag_product_list, FILENAME_ABC, False)

if tag_product_haw != "":
	tag_product_list = "<product_list>\n%s</product_list>" % tag_product_haw.encode('utf-8')
	writeOutputFile(tag_product_list, FILENAME_HAW, False)

connection_sql.close()
connection_erp.close()
printForDebug("--- %s seconds ---" % (time.time() - start_time), FLG_DEBUG)