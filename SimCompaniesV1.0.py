import openpyxl
import requests
import json
import math as Math
from datetime import datetime
import time

######## INIT API ########

URL_ency = "https://www.simcompanies.com/api/v4/fr/0/encyclopedia/resources/1/"

######## INIT SIM COMPANIES ########

#               {"name of product' : {"URL" : "3/", "price" : 0, "cout MP" : 0, "cout transport" : 0}
speed_prod = 1.01
speed_sell = 1-0.03
costAdministration = 0.1118
products_dic = {}
            
buildings_prod = {}

######## INIT EXCEL #########
dest_filename = "SimCompaniesTest.xlsx"
wb = openpyxl.Workbook()
wb = openpyxl.load_workbook(filename = dest_filename)


######## FUNCTIONS #########

def update_ressource_list():
    global products_dic
    print("Update list of products...")
    URL_ressource = "https://www.simcompanies.com/api/v4/en/0/encyclopedia/resources/"
    reponse = requests.get(URL_ressource)
    contenu = reponse.json()#récupère la réponse en chaine de caractère
    for k in contenu:
        x = json.dumps(k)#récupère la première chaine de contenu et la converti en str
        x = json.loads(x) #converti la chaine x en dictionnaire    
        products_dic.update({x["name"]:{"URL" : x["db_letter"], "price" : 0, "ROI SELL IA" : 0, "ROI SELL market" : 0, "price IA" : 0,"cost admin sell IA" : 0,"cost admin sell market" : 0, "cost MP" : 0, "cost MO prod" : 0, "cost MO sell" : 0, "transport" : 0, "cost transport" : 0, "prod /h" : 0, "sell /h" : 0,"benef /h sell IA" : 0, "benef /h sell market" : 0,"best price IA" : 0, "cost wages sell IA/h" : 0}})
    #print(products_dic)
    
def update_price():
    URL_market = "https://www.simcompanies.com/api/v3/market/0/"
    global products_dic
    print("Update price of each product...")
    counter = 0
    for key in products_dic :
        counter += 1
        URL = URL_market + (str)(products_dic[key]['URL']) +'/'
        requests.get("https://www.simcompanies.com/api/v2/companies/me/")
        requests.get("https://www.simcompanies.com/api/realms/")
        requests.get("https://www.simcompanies.com/api/v2/companies/me/")
        requests.get("https://www.simcompanies.com/api/v1/players/me/companies/")
        requests.get("https://www.simcompanies.com/api/v3/contracts-incoming/0/me/")
        requests.get("https://www.simcompanies.com/api/v2/market-ticker/0/2024-04-18T20:35:35.784Z/")
        reponse = requests.get(URL)
        contenu = reponse.json()#récupère la réponse en chaine de caractère
        print(contenu)
        print("\n")
        print("produit qui plante :",key)
        if counter == 10 :
            time.sleep(30)
            counter = 0
        #si il n'y a aucune offre on ne met pas le prix à jour 
        try :
            if len(contenu)>0:
                x = json.dumps(contenu[0])#récupère la première chaine de contenu et la converti en str
                x = json.loads(x) #converti la chaine x en dictionnaire
                products_dic[key]["price"]=x["price"]
        except IndexError:
            #print("Aucun produit "+ key + " n'est vendu en bourse")
            pass

###

def calculate_cost_products():
    global products_dic
    print("Calculate the cost of raw materials, transport, for each product...")
    print("Looking for the better price to sell in the market and calculate the bigger profit...")
    URL_ressources = "https://www.simcompanies.com/api/v4/en/0/encyclopedia/resources/2/"

    for key in products_dic :
        cost = 0
        URL = URL_ressources + (str)(products_dic[key]['URL'])
        reponse = requests.get(URL)
        contenu = reponse.json()#récupère la réponse en chaine de caractère
        x = json.dumps(contenu)#récupère la première chaine de contenu et la converti en str
        x = json.loads(x) #converti la chaine x en dictionnaire
        produced_hour(x)
        calculate_cost_transport(x)
        calculate_cost_MP(x)
        update_price_IA(x)
        calculate_cost_sell()
        benef_selling_IA(x)

###

def calculate_cost_MP(x):
    global products_dic
    cost = 0
    if(x["producedFrom"]!=[]):
        for resource in x["producedFrom"]:
            cost = cost + products_dic[resource["resource"]["name"]]["price"] * resource["amount"]
        products_dic[x["name"]]["cost MP"] = cost

###
        
def calculate_cost_transport(x):
    global products_dic
    products_dic[x["name"]]["transport"]=x["transportation"]
    products_dic[x["name"]]["cost transport"]=products_dic[x["name"]]["transport"] * products_dic["Transport"]["price"]

###

def update_price_IA(x):
    global products_dic, speed_prod
    if(x["retailable"]== True):
        try :
            products_dic[x["name"]]["price IA"]=(float)(x["averageRetailPrice"])
        except TypeError:
            #print("Aucun prix moyen n'a été trouvé en vente IA pour " + x["name"])
            pass
    else:
        products_dic[x["name"]]["price IA"]=-1

###

def produced_hour(x):
    global products_dic, speed_prod
    products_dic[x["name"]]["prod /h"]=(float)(x["producedAnHour"]) * speed_prod

###

def benef_selling_IA(x):
    global products_dic, speed_sell, costAdministration

    if(x["retailable"]== True):
        try:
            bestBenefHour=0
            bestCostMoSell = 0
            price = 0.6*x["averageRetailPrice"]
            end_price = 1.4*x["averageRetailPrice"]
            best_price = price
            saturation = x["marketSaturation"]
            amount=1
            bestTime=eval((x["retailModeling"]))
            while(price<end_price):
                time_selling =eval((x["retailModeling"]))
                if(time_selling != 0):
                    products_selling_hour = (3600./(time_selling*speed_sell))
                    costMOsell = (float)(products_dic[x["name"]]["cost wages sell IA/h"])/(products_selling_hour)
                    benefHour=products_selling_hour*(price-products_dic[x["name"]]["price"]-costMOsell-costAdministration*costMOsell)
                    if(benefHour>bestBenefHour):
                        bestBenefHour = benefHour
                        bestTime=time_selling*speed_sell
                        best_price = price
                        bestCostMoSell = costMOsell
                        bestCostAdmin = costMOsell*costAdministration
                price = price + 0.01*x["averageRetailPrice"]
            if(bestTime != 0):
                products_dic[x["name"]]["sell /h"]=3600./bestTime
                products_dic[x["name"]]["benef /h sell IA"] = bestBenefHour
                products_dic[x["name"]]["best price IA"] = best_price
                products_dic[x["name"]]["cost MO sell"]= bestCostMoSell
                products_dic[x["name"]]["cost admin sell IA"]= bestCostMoSell * costAdministration
        except TypeError:
            #print("Le produit n'a pas de calcul de modelisation de besoin de type x["retailModeling"]")
            pass           

###
        
def benef_selling_market():
    global products_dic
    for resource in products_dic:
        benefHour=products_dic[resource]["prod /h"]*((products_dic[resource]["price"])*0.97-products_dic[resource]["cost MP"]-products_dic[resource]["cost MO prod"]-costAdministration*products_dic[resource]["cost MO prod"]-products_dic[resource]["cost transport"])
        products_dic[resource]["benef /h sell market"]=benefHour
        products_dic[resource]["cost admin sell market"]=costAdministration*products_dic[resource]["cost MO prod"]
    
###
 
def update_buildings_prod():
    global buildings_prod

    print("Update information of buildings production")
    URL_buildings = "https://www.simcompanies.com/api/v3/0/buildings/1/"
    reponse = requests.get(URL_buildings)
    contenu = reponse.json()#récupère la réponse en chaine de caractère
    for k in contenu:
        x = json.dumps(k)#récupère la première chaine de contenu et la converti en str
        x = json.loads(x) #converti la chaine x en dictionnaire
        buildings_prod.update({x["name"]:{"salaire" : x["wages"] ,  "cout de construction" : x["cost"], "production" : [], "vente" : [] }})

        try:
            for i in range(len(x["production"])):
                buildings_prod[x["name"]]["production"].append(x["production"][i]["resource"]["name"])                
        except KeyError:
            #print("Le " + x["name"] + " ne produit rien, il vend seulement !")
            pass

        try:
            for i in range(len(x["retail"])):
                buildings_prod[x["name"]]["vente"].append(x["retail"][i]["resource"]["name"])
        except KeyError:
            #print("Le " + x["name"] + " ne vend rien, il produit seulement !")
            pass
        
###

def calculate_cost_prod():
    global buildigs_prod, products_dic
    
    for resource in products_dic:
        for building in buildings_prod:
            if(resource in buildings_prod[building]["production"] and (products_dic[resource]["prod /h"]) != 0 ):
                #print("Le bâtiment " + building + " produit " + resource)
                #print(resource + " prod")
                products_dic[resource]["cost MO prod"] = (float)(buildings_prod[building]["salaire"])/(products_dic[resource]["prod /h"])

###

def calculate_cost_sell():
    global buildigs_prod, products_dic
    
    for resource in products_dic:
        for building in buildings_prod:
            if(resource in buildings_prod[building]["vente"]):
                #print(resource + " vente")
                #products_dic[resource]["cost MO sell"] = (float)(buildings_prod[building]["salaire"])/(products_dic[resource]["sell /h"])
                products_dic[resource]["cost wages sell IA/h"] = (float)(buildings_prod[building]["salaire"])
            
###

def calculate_ROI():
    global buildigs_prod, products_dic
    
    for resource in products_dic:
        for building in buildings_prod:
            if(resource in buildings_prod[building]["vente"] and products_dic[resource]["benef /h sell IA"] != 0):
                products_dic[resource]["ROI SELL IA"] = ((products_dic[resource]["benef /h sell IA"]*24)/buildings_prod[building]["cout de construction"])*100
            if(resource in buildings_prod[building]["production"] and products_dic[resource]["benef /h sell market"] != 0):
                products_dic[resource]["ROI SELL market"] = ((products_dic[resource]["benef /h sell market"]*24)/buildings_prod[building]["cout de construction"])*100
###

def products_with_most_ROI():
    global products_dic
    ROI = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]
    name = ['','','','','','','','','','']
    for resource in products_dic:
        for i in range(0,10):
            if(products_dic[resource]["ROI SELL market"]>ROI[i]):
                for j in range(9,i,-1):
                    ROI[j]=ROI[j-1]
                    name[j]=name[j-1]
                ROI[i]=products_dic[resource]["ROI SELL market"]
                name[i]=resource
                break
    print()
    print("ROI market ")
    for i in range(0,10):              
        print(name[i] + " : " + (str)(round(ROI[i],2))+ "%")

###

def products_with_most_ROI_IA():
    global products_dic
    ROI = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]
    name = ['','','','','','','','','','']
    for resource in products_dic:
        if(products_dic[resource]["ROI SELL IA"]>0):
            for i in range(0,10):
                if(products_dic[resource]["ROI SELL IA"]>ROI[i]):
                    for j in range(9,i,-1):
                        ROI[j]=ROI[j-1]
                        name[j]=name[j-1]
                    ROI[i]=products_dic[resource]["ROI SELL IA"]
                    name[i]=resource
                    break
                
    print()
    print("ROI SELL IA ")
    for i in range(0,10):              
        print(name[i] + " : " + (str)(round(ROI[i],2))+ "%")

###

def products_with_most_profits(cmd):
    global products_dic
    
    benef = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]
    name = ['','','','','','','','','','']
    for resource in products_dic:
        if(products_dic[resource][cmd]>0):
            for i in range(0,10):
                if(products_dic[resource][cmd]>benef[i]):
                    for j in range(9,i,-1):
                        benef[j]=benef[j-1]
                        name[j]=name[j-1]
                    benef[i]=products_dic[resource][cmd]
                    name[i]=resource
                    break
                
    print()
    print("BENEF "  + cmd)
    for i in range(0,10):              
        print(name[i] + " : " + (str)(round(benef[i],2))+ "$")
            
        
######### END FUNCTIONS #########
        
######### MAIN #########

#update_buildings_prod()
update_ressource_list()
update_price()
"""
calculate_cost_products()
calculate_cost_prod()
benef_selling_market()
calculate_ROI()

print(buildings_prod)
print(products_dic)

print(str(datetime.now()))
products_with_most_profits("benef /h sell IA")
products_with_most_profits("benef /h sell market")
products_with_most_ROI()
products_with_most_ROI_IA()




#####try a save on XLSX #####

"""
try:
    (wb.save(filename = dest_filename))
except PermissionError:
    print("ERROR : You file xlsx is open. Please close it.")

