#Doğrulama Kodu
import requests
from bs4 import BeautifulSoup
url = "https://docs.google.com/spreadsheets/d/1AP9EFAOthh5gsHjBCDHoUMhpef4MSxYg6wBN0ndTcnA/edit#gid=0"
response = requests.get(url)
html_content = response.content
soup = BeautifulSoup(html_content, "html.parser")
first_cell = soup.find("td", {"class": "s2"}).text.strip()
if first_cell != "Aktif":
    exit()
first_cell = soup.find("td", {"class": "s1"}).text.strip()
print(first_cell)


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager



print("Excelin İlk Sütununa Sadece Tamamlandı Yapılacak Linkler Listelenir")
print("Excelin Adı Tamamlandı Yapılacak Olması Lazım")



# ChromeDriver'ı en son sürümüyle otomatik olarak indirip kullan
driver = webdriver.Chrome(ChromeDriverManager().install())

# Giriş yap
login_url = "https://task.haydigiy.com/kullanici-giris/?ReturnUrl=%2Fadmin"
driver.get(login_url)

email_input = driver.find_element("id", "EmailOrPhone")
email_input.send_keys("mustafa_kod@haydigiy.com")

password_input = driver.find_element("id", "Password")
password_input.send_keys("123456")
password_input.send_keys(Keys.RETURN)

# Excel dosyasını oku
excel_file_path = "Tamamlandı Yapılacak.xlsx"
df = pd.read_excel(excel_file_path, header=None, names=["Link"])

# Her bir link için işlemi gerçekleştir
for index, row in df.iterrows():
    link = row["Link"]
    
    # Linki aç
    driver.get(link)

    # JavaScript kodunu çalıştır
    js_code = """
    var selectElement = document.querySelector("#OrderStatusId");
    selectElement.value = "3"; // "Tamamlandı" seçeneği
    var event = new Event("change", { bubbles: true });
    selectElement.dispatchEvent(event);
    
    var confirmButtonElement = document.querySelector("#btnSaveOrderStatus-action-confirmation-submit-button");
    confirmButtonElement.click();
    """
    driver.execute_script(js_code)

# Browser'ı kapat
driver.quit()