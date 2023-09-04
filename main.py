from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
import pandas as pd
import json
import time
import os

def Click(driver, xpath, t = 1):
	try:
		time.sleep(1)
		WebDriverWait(driver, t).until(expected_conditions.presence_of_element_located((By.XPATH, xpath)))
		driver.find_element(By.XPATH, xpath).click()	
		return 1
	except:
		return

def Input(driver, xpath, texto, t = 1, Status=True):
	WebDriverWait(driver, t).until(expected_conditions.element_to_be_clickable((By.XPATH, xpath)))
	if Status:
		driver.find_element(By.XPATH, xpath).click()
		driver.find_element(By.XPATH, xpath).send_keys(Keys.CONTROL + 'a',Keys.BACKSPACE,Keys.BACKSPACE)
	driver.find_element(By.XPATH, xpath).send_keys(texto,Keys.TAB)
	if Status:
		driver.find_element(By.XPATH, '//*[@id="content"]/header/div/div/h1').click()


def Team(driver,xpath,texto, t=1):
	try:
		WebDriverWait(driver, t).until(expected_conditions.element_to_be_clickable((By.XPATH,xpath )))
		driver.find_element(By.XPATH, xpath).send_keys(texto)
		try:
			WebDriverWait(driver, t).until(expected_conditions.element_to_be_clickable((By.XPATH, '//div/div/div/span/b[contains(text(),"'+Tablero+'")]')))
			driver.find_element(By.XPATH, '//div/div/div/span/b[contains(text(),"'+Tablero+'")]').click()
			return 1
		except:
			return
	except:
		return

with open('config.json') as f:
	config = json.load(f)

URLJira = config["URLJira"]
RutaUserData = config["RutaUserData"]
ProfileBBVA = config["ProfileBBVA"]
MVP = config["MVP"]

if os.path.isfile("./Input.xlsx"):
	Df = pd.read_excel('./Input.xlsx', sheet_name='TAREAS', dtype = 'object', usecols="B:C",nrows=1)
	TablaDf = pd.read_excel('./Input.xlsx', sheet_name='TAREAS', dtype = 'object', usecols="A:H",header=3).dropna(how="all")
	Tablero = Df['TABLERO'][0]
	options = webdriver.ChromeOptions()
	options.add_argument(f'--user-data-dir={RutaUserData}')
	options.add_argument(f'--profile-directory={ProfileBBVA}')
	driver = webdriver.Chrome(service=Service(), options=options)
	driver.maximize_window()
	driver.get('https://www.google.com')
	for index, row in TablaDf.iterrows():
		if row['HUT'] != row['HUT']:
			driver.get(f'{URLJira}')
			Input(driver, '//*[@id="project-field"]', 'Peru App Datio (PAD3)', 5)
			if '[Equipo]' in row['CODIGO'] and 'Control M' in row['CODIGO']:
				Input(driver, '//*[@id="issuetype-field"]', 'Dependency', 5)
			else:
				Input(driver, '//*[@id="issuetype-field"]', 'Story', 5)
				if Click(driver, '//*[@id="issue-create-submit"]',5) is not None:

					Team(driver,'//label[text()="Team Backlog "]/following-sibling::div//input', Tablero, 5)
					if '[Equipo]' in row['CODIGO'] and 'Control M' in row['CODIGO']:
						Team(driver,'//label[text()="Receptor Team "]/following-sibling::div//input', Tablero, 5)
						Input(driver, '//*[@id="labels-textarea"]', 'release', 5, False)
						Input(driver, '//*[@id="labels-textarea"]', 'ReleaseMallasDatio', 5, False)
						Input(driver, '//*[@id="customfield_10267-textarea"]', 'AppsInternos', 5, False)
						Input(driver, '//*[@id="customfield_10267-textarea"]', 'Datio', 5, False)
						Click(driver, '//*[@id="customfield_10270"]/div[2]/label')
						Click(driver, '//*[@id="customfield_18001"]/option[14]')
						if MVP is not None or MVP != "":
							text = f'Como equipo declaramos que el siguiente Pase está listo para transitar por las etapas de Certificación QA y Pase a Producción, y la documentación adjunta corresponde con el MVP {MVP} así como las Historias de Usuario enlazadas a este pase.'
							Input(driver, '//*[@id="customfield_10260"]', text, 5)

					else:
						if '[Equipo]' in row['CODIGO'] and 'Validar PR' in row['CODIGO']:
							Input(driver, '//*[@id="labels-textarea"]', 'ReleasePRDatio', 5)
							Click(driver, '//*[@id="customfield_10260-wiki-edit"]/nav/div/div/ul/li[2]/button' )
							Input(driver, '//*[@id="customfield_10260"]', 'Desarrollo según los Lineamientos del Equipo de DQA.', 5)
							if 'SmartCleaner' in row['CODIGO']:
								Input(driver, '//*[@id="labels-textarea"]', 'SmartCleaner', 5)
							if 'Ingesta' in row['CODIGO']:
								Input(driver, '//*[@id="customfield_10267-textarea"]', 'Ingesta', 5)
							if 'Hammurabi' in row['CODIGO']:
								Input(driver, '//*[@id="customfield_10267-textarea"]', 'Hammurabi', 5)
							Click(driver, '//*[@id="customfield_10270"]/div[2]/label')
							Click(driver, '//*[@id="customfield_18001"]/option[14]')
						else:
							Click(driver, '//*[@id="assign-to-me-trigger"]')
							Click(driver, '//*[@id="customfield_10270"]/div[1]/label')
							time.sleep(1)
							Click(driver, '//*[@id="customfield_10270"]/div[2]/label')
							time.sleep(1)
							Click(driver, '//*[@id="customfield_10270"]/div[1]/label')
							Click(driver, '//*[@id="customfield_10270"]/div[1]/label')
					Code = row['CODIGO'].replace('[Equipo]',f'[{Tablero}]').replace('[UUAA]',row['UUAA']).replace('[fuente]',row['FUENTE']).split(' - ', 1)
					Input(driver, '//*[@id="summary"]', Code[1], 5)
					Input(driver, '//label[text()="Feature Link"]/following-sibling::div/div/input', row['FEATURE'], 5)
					time.sleep(1)
					Click(driver,'//ul[@id="suggestions"]/li[a/span/em[contains(text(),"'+row['FEATURE']+'")]]')
					Input(driver, '//*[@id="labels-textarea"]', 'P-'+Code[0], 5, False)
					Input(driver, '//*[@id="labels-textarea"]', 'F-'+row['FOLIO'], 5, False)
					Input(driver, '//*[@id="labels-textarea"]', 'ID-'+str(row['ID']), 5, False)
					Click(driver, '//*[@id="description-wiki-edit"]/nav/div/div/ul/li[2]/button' )
					Input(driver, '//*[@id="description"]', row['FUENTE'], 5)
					
					time.sleep(1)
					
					Click(driver,'//*[@id="issue-create-submit"]')
					WebDriverWait(driver, 10).until(expected_conditions.presence_of_element_located((By.XPATH, '//*[@id="key-val"]')))
					print(driver.find_element(By.XPATH, '//*[@id="key-val"]').get_attribute("href"))
					time.sleep(1)
