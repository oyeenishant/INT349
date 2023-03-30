from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import re
import time
import csv
import pandas as pd
from selenium.webdriver.common.by import By
driver =webdriver.Chrome()

# Define the URL of the Amazon product
url = "https://amzn.eu/d/7fNyfUW"

# Initialize the Chrome driver
driver = webdriver.Chrome()

# Navigate to the Amazon product page and open the reviews section
driver.get(url)
time.sleep(5)  # Wait for the page to load
reviews_button = driver.find_element(By.XPATH, '//*[@id="reviews-medley-footer"]/div[2]/a')
time.sleep(2)
reviews_button.click()
time.sleep(1)  # Wait for the reviews section to load

# Scroll down to load all reviews
while True:
    try:
        load_more_button = driver.find_element(By.XPATH, '//*[@id="cm_cr-pagination_bar"]/ul/li[last()]/a')
        load_more_button.click()  # Click the "Load more" button to load more reviews
        time.sleep(2)  # Wait for the new reviews to load
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.END)  # Scroll to the end of the page again
    except:
        break  # No more "Load more" button found

time.sleep(5)

# Extract the reviews from the page source
reviews = driver.find_elements(By.XPATH, '//div[@data-hook="review"]')
review_list = []
for review in reviews: 
    text_element = review.find_element(By.XPATH, './/span[@data-hook="review-body"]')
    text = text_element.text    
    review_list.append({'text': text})

# Save the reviews to a CSV file
with open('amazon_reviews.csv', 'w', encoding='utf-8', newline='') as csvfile:
    fieldnames = ['text']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    for review in review_list:
        writer.writerow(review)
# Close the browser
driver.quit()

import pandas as pd
import re

# Read the CSV file into a DataFrame
df = pd.read_csv('amazon_reviews.csv')
# Write the DataFrame to an Excel file
df.to_excel('output_file.xlsx', index=False)
# Read in Excel file
df = pd.read_excel('output_file.xlsx')


import openai
from pyChatGPT import ChatGPT
session_token='eyJhbGciOiJkaXIiLCJlbmMiOiJBMjU2R0NNIn0..vZvG6QNEqjUhG5lT.Fh_F2DQ9GZ4Fu6lZ_cs9qZpa8DEI1d0JMM_Iu8qGQDveLlkAwPv86kGiu_zRyqViRyPyXMUAppDPaO4s7-wy9ev6P1YImZ-TAoHI_rvsD0dHNwM1NEhP_Fa5OL8vKzFunHOgSMGFKW2wgJJKFxyLvg8gRdewygrORkFWNmHgZOJ2TMHNwLl478l5_QHvg6FG3c3xjtIX0IoIeL26E-5C2YsGN1FpwOETrFgNfQy9DD1Ag2Lxargd3mmi0mOS6Vf5bHO5KcztcJ7LstA0HGmmu4E8Ct6e0i-oJnv16em8fpY24EqKYYaGxZ3jkLptpEkbigJNe9OH57esYIOPkGHoBV-qLTS5BgiemyIWLwKBTB8-HumMopLc0p95zjJpAbpz9aPRZCPBf72h76Ok68ZMvHC-n_RwTsMoIIykR1hLtQv8vvRJnUCGNtyPeROAupF94x4O6RCFiQXVqFj5YNxc2gbNKEiEPgY4iyhffE9Mxe7R1rzySJBiNSiCSNfPU3IVfIgfCR7grRmQSHb5mzR1GLT3h6rwO1tFLU2HCz4prLe0ZXrCW_p_P1GB8UkF0tjvVYaFg6jsOsKGBEuLA6a1wxk5J1HzjPR6rBTzNEm2HQp4Kzm3tY7pCXXz1d13c9L0YTuG5kK8uRNJCWIlO7P61te70i0P2BYrBwDljCRNbC8w0R9vAb7rgPDe2RBDztx8HmXHeNRfbzXwJ2_Mk_EW7CvLnVnhi4TQSl8MKtSFDYs3YT355QUgSXroJFRe-BAz_2fz35ztM3pj183iYd5kRTp3Q3WYhuFgwxCNQ1vJ55vTTk675RQ4Uj-m3SvbZQ_FMg9D7LuF6HQftxM3kXTQZwt0bbkGTLeKM0iWRNh3xFlzxYsZhW-kjC15riiud7ZAdpDwhBBiim_oo7APk5a6hheEVDc9f6PLSN_kS4xqyGtcwufKGUM6zFhpCNrKrDLLfWb1_sLrmNTu_2WRnH80nYGRVhkJfeUi72kvH5oTakBrqweet5eLr0QALNdSm3Gd5lOxvZ-KJwNk5zCxGafrEkUP_io9olvV4GWwZGocuHe742xmmVr6p3tJOgIGK4c3ti4Anm_CKPEfIGdRbZCObNkIKNXQKGIixyCEdM6v-9vBz2qi1mmA4lTclwQxQqGEZ65JeBHzG1yaZfYSIVZi77MNXMYyowczZZbSU5xBTb-4WCJ-tzxpIEy5-OXdAhCs6mjMOvYxSHKy2QOaQTwWj-JNgHUGzdZ2HcjTGcRdc8IorWWDE6O3mkWxw7OeVnAVP3STj187dAguF8emblvPTRAJEe1thirf12ZIamjY4o-LRXARmuicrH4x5wVD4bZxZoX5DJc3E4mVwo59z93jqYYmRFuCB8BeKWVoonQHWorUq8O-inroce-MdcqfliiXyM_Xwajtt66QNOHvL6cAT94nM-c2dRx4QozTRChyrabp1d6M69MMRylYBqcSvLZFRVg8bCM0IAcMnpnF7VRmlFModql4jUppoacJ75VWS6fqFaRWVFwzJXSvPIRhfHhTEvX6vmVw5Jn64b2WbQ8TwOhnW3-2uSh8LZp9_T-Qsl2JBnQnzoRyVzoKl88Wh1VlTR0U5vkS3JsukJjRIxYpbViw2zofthHPfP0sih-sYK5XKbJAut36qZ0cl8yIRKFiNTQBubcg1HvQtqdO2Olo4LT4oEK3BnoE93sJ3KZ0_mfcagxtr1v1qRiS3A1S94AtmlUxR6s-OhOqEnqCZWXSkv0v9DMn8EmpR0UusBam4sWJoFV5I4Bp39KYZLPZlXnt-s30FfDkaAqofBJRDWi9HvCqubENOd_EXm45buDoJbYB6lpQRfY56PeBdC2E03bqa6dM9sreZBdbo507bw7BYhSZ_6Qsc7Cdie32Lyw7ejjO8nxJXkSpR7FMnZ4Zz6po-W_1DR8ZP7LxVs-aSCwnJb12egT2iPXMlgAWvgtuRhXmcgp1Y_UhFMNMhyl1F15ATbOLc-JnNG8PczxFm7OLR0pf9GuYZYvStxOVwb4SuxuLE9OtfbQR0RiOX9gMZ8SA3TVYFHl5F32W9FLr1K8hHYDapuFt6e0s6NxUyrQOGw7beRUl4J90z2BpuUhGmc0IKc9fOAGhEfO9JZAUULkTLAjL9W-mN-14SD0fQZE9JsQhuxVsG4zZmwFd1bqjKaTK5Gy3Vu9bON3LP4bgMMrSf2zQpW1eoxWyaWcDcLOEzv-ZQfC3dht9RXlqFQYaaJiqGz2TGgpdrpAR5vX3Sxvylj4UmOf1-ZSl00tLDD-SGj5lAm5KuQVOdXgQXl-ehWWWs3FL65K37WnjIy_mtmDRNvMYtfEaGVz9s-unKi_DaMIUPBA47tujsZ2upNNGNKJ7agP0I2sYaZ4j5nqX5-gw_9pD4haPXFyY2OCEuTvonttCLhv_GP9a-UvY5nug9on4zViB9DKN50qv9AZEQaSVoCNyJzuMFzNv8oc9ISjW4RJQf9OWrU1OA39ApE07QPvc-F6YnToLB7Y.hSnQf05INTVEYp6xew5afQ'
api=ChatGPT(session_token)
import time
time.sleep(10)
# Iterate over each row in the Excel file
for index, row in df.iterrows():
    # Get the value of the 'input' cell in this row
    input_text = row['text']
    time.sleep(3)
    resp=api.send_message("Give me answer in one word, this is postive or negative or netural sentence "+input_text)
    time.sleep(3)
    output_text=resp['message']
    print(output_text)
    text = output_text
    clean_text = re.sub(r'\.', '', text)
    print(clean_text)
    # Write the generated text to the 'output' cell in this row
    df.at[index, 'Sentiments'] = clean_text
    api.refresh_chat_page()
    api.reset_conversation()

# Write the updated dataframe back to the same Excel file
df.to_excel('output_file.xlsx', index=False)
