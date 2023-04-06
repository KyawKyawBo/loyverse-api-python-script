import requests # http request အတွက် သုံးသည့် library ဖြစ်ပါသည်။ မရှိသေးပါက pip ကို သုံး၍ ထည့်သွင်းရပါမည်။
import openpyxl # Excel ဖိုင်များကို handle လုပ်ရန် အတွက် သုံးသည့် library ဖြစ်ပါသည်။ မရှိသေးပါက pip ကို သုံး၍ ထည့်သွင်းရပါမည်။

# Giving Excel File Name , မိမိ သိမ်းတဲ့အခါ ပေးလို့သည့် Excel ဖိုင်အမည်ကို ပေးပါ။
file_name = "<TYPE_YOUR_EXCEL_FILE_NAME>"

# Loyverse API endpoint
url = "https://api.loyverse.com/v1.0/items"

""" Loyverse API access token , ကိုယ့် Loyverse Store မှာ Integration ဝယ်ထားရင် Access Token ထုတ်လို့ရပါတယ်၊
အသေးစိတ်သိချင်ရင် ဒီလင့်ခ်မှာ ကြည့်ပါ။ https://help.loyverse.com/help/loyverse-api """
access_token = "<Your_Loyverse_API_Access_Token>"

# Request headers
headers = {
    "Authorization": "Bearer " + access_token,
    "Content-Type": "application/json"
}

# Request parameters
params = {
    "limit": 250,
    'cursor': '<Cursor_ID>' 
  """ if you pull from the start, just make this line a comment. if not, put your Cursor_ID
  ပထမဆုံး အကြိမ် pull လုပ်တာဆိုရင်တော့ ဒီ field က မလိုပါဘူး။ ပထမ အသုတ် တစ်ခုကို pull လုပ်လိုက်တဲ့အခါ
  JSON မှာ cursor id တခါတည်း ပါလာပြီး ပထမ အသုတ်က ဟာတွေကို ထပ်မဆွဲတော့ဘဲ ဒုတိယ အသုတ်အနေနဲ့ နောက်ထပ်
  ကျန်နေသေးတဲ့ဟာတွေကို ဆက်ဆွဲချင်တဲ့အခါမှာ ဒီ field ကို အသုံးပြုရမှာ ဖြစ်ပါတယ်။ """
}


# Send GET request to Loyverse API , Loyverse API ဆီ GET request ပို့မယ်
response = requests.get(url, headers=headers, params=params)


print(response.status_code)
print(response.json())


# Check if request was successful , request အောင်မြင်မှု ရှိ/မရှိ စစ်ဆေးမယ်
if response.status_code != 200:
    print("Error: Failed to get items from Loyverse API")
    exit()

# Parse response JSON and create list of items , ရလာတဲ့ JSON ကို parse လုပ်ပြီး item list ဖန်တီးမယ်
items = response.json()['items']

# Create new Excel workbook and worksheet , Excel workbook အသစ်တစ်ခု နဲ့ worksheet တစ်ခု ဖန်တီးမယ်
workbook = openpyxl.Workbook()
worksheet = workbook.active

# Write headers to worksheet , worksheet မှာ header တွေ ထည့်မယ်
headers = ['Item ID', 'Handle', 'Item Name', 'Description', 'Reference ID', 'Category ID', 'Track Stock', 'Sold By Weight', 'Is Composite', 'Use Production', 'Primary Supplier ID', 'Form', 'Color', 'Image URL', 'Option 1 Name', 'Option 2 Name', 'Option 3 Name', 'Created At', 'Updated At', 'Deleted At', 'Variant ID', 'SKU', 'Reference Variant ID', 'Option 1 Value', 'Option 2 Value', 'Option 3 Value', 'Barcode', 'Cost', 'Purchase Cost', 'Default Pricing Type', 'Default Price', 'Store ID', 'Pricing Type', 'Price', 'Available for Sale', 'Optimal Stock', 'Low Stock']
worksheet.append(headers)

# Loop through items and write data to worksheet , item ဒေတာတွေကို loop လုပ်ပြီး worksheet ထဲထည့်မယ်
for item in items:
    # Write item data to worksheet
    row_data = [item['id'], item['handle'], item['item_name'], item['description'], item['reference_id'], item['category_id'], item['track_stock'], item['sold_by_weight'], item['is_composite'], item['use_production'], item['primary_supplier_id'], item['form'], item['color'], item['image_url'], item['option1_name'], item['option2_name'], item['option3_name'], item['created_at'], item['updated_at'], item['deleted_at']]

    for variant in item['variants']:
        # Write variant data to worksheet
        row_data.extend([item['id'], variant['sku'], variant['reference_variant_id'], variant['option1_value'], variant['option2_value'], variant['option3_value'], variant['barcode'], variant['cost'], variant['purchase_cost'], variant['default_pricing_type'], variant['default_price'], variant['stores'][0]['store_id'], variant['stores'][0]['pricing_type'], variant['stores'][0]['price'], variant['stores'][0]['available_for_sale'], variant['stores'][0]['optimal_stock'], variant['stores'][0]['low_stock']])
        
        # Write component data to worksheet
        for component in item['components']:
            row_data.extend([component['variant_id'], component['quantity']])
        
        worksheet.append(row_data)
        row_data = [''] * len(headers)

# Save workbook , workbook ကို စစ်မယ်
workbook.save(filename=file_name)

print("Successfully pulled items from Loyverse API and saved to Excel file.")

