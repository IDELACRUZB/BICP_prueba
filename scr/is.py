import json

pathjson1392 = 'parametro_campana/bicp1392.json'
with open(pathjson1392, 'r') as json1392:
    parametros1392 = json.load(json1392)

for campana, parametros in parametros1392.items():
    nombreCampa = campana
    bpoXpath = parametros["bpoXpath"]
    campaignXpath = parametros["campaignXpath"]

    print(nombreCampa)
    print(bpoXpath)
    print(campaignXpath)
    #print(xpathAgentWorkgroup)
        
    print(10*"=")
