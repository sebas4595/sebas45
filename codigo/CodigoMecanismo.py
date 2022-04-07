#codigo de mecanismo complementario de la tercera subasta de renovables
from Subasta import *

datos1 = getcwd() + r"\Adjudicacion_Subasta_CLPE-01-2021.XLSX" 
DO = 5520000

resultado1 = AlgoritmoOptimizador(datos1)

b = resultado1[0] #adjudicacion
c = resultado1[1] #venta maxima
d = resultado1[2] #TEAMC (comprador)
e = resultado1[3] #TEAMC (vendedores)

print(b,c,d,e)

if b < 0.95*DO and d>0.05*DO and b < 0.7*DO:
	print("se aplica el mecanismo de activacion 2")
	
	#Lectura de datos de excel
	t = []
	L0 = []
	L1 = []
	L2 = []
	L3 = []
	L4 = []
	L5 = ['nombre','ID_oferta','compra_max','precio','ordenLlegada','EAMC']
	L6 = ['nombre','ID_oferta','bloque','venta_max','venta_min','precio','simultanea','excluyente','dependiente','ordenLlegada']
	L7 = []
	L8 = ["Comprador", "ID_oferta", "Asignacion_de_compra", 'EAMC']
	L9 = ["Vendedor", "ID_oferta", "Bloque", "Asignacion_de_venta"]
	modelo = ConcreteModel("SubastaRenovables")
	datos = getcwd() + r"\Adjudicacion_Subasta_CLPE-01-2021.XLSX" 
	xlFile = datos
	xl_datos = pd.ExcelFile(xlFile)
	precio_compra = xl_datos.parse("compradores").set_index(["ID_oferta"])
	precio_venta = xl_datos.parse("vendedores").set_index(["ID_oferta"])
	precio_compra1 = xl_datos.parse("resultados_compradores").set_index(["ID_oferta"])
	precio_venta1 = xl_datos.parse("resultados_vendedores").set_index(["ID_oferta"])
	modelo.compra = Set(initialize=precio_compra.index)
	modelo.venta = Set(initialize=precio_venta.index)		
	#---------------------------------------------------------
	
	for i in modelo.compra:
		L1.append(precio_compra.compra_max[i]-precio_compra1.Asignacion_de_compra[i])
		L0.append(precio_compra.nombre[i])
		t.append(precio_compra1.Asignacion_de_compra[i])
	for i in modelo.venta:
		L2.append(precio_venta.venta_max[i]-precio_venta1.Asignacion_de_venta[i])
		L7.append(precio_venta.nombre[i])
	
	#print(L1,L2)
	
	for i in range(len(L1)):
		if L1[i] != 0:
			L4.append(190)
	#print(L3)
	archivo = xlsxwriter.Workbook('resultados_mecanismo.xlsx')
	hoja1 = archivo.add_worksheet('compradores')
	hoja2 = archivo.add_worksheet('vendedores')
	hoja3 = archivo.add_worksheet('resultados_compradores')
	hoja4 = archivo.add_worksheet('resultados_vendedores')
	c = 0
	for i in range(len(L1)):
		if c == 0:
			hoja1.write(0,0,L5[i])
			hoja1.write(0,1,L5[i+1])
			hoja1.write(0,2,L5[i+2])
			hoja1.write(0,3,L5[i+3])
			hoja1.write(0,4,L5[i+4])
			hoja1.write(0,5,L5[i+5])
			hoja3.write(0,0,L8[i])
			hoja3.write(0,1,L8[i+1])
			hoja3.write(0,2,L8[i+2])
			hoja3.write(0,3,L8[i+3])
			hoja4.write(0,1,L9[i])
			hoja4.write(0,2,L9[i+1])
			hoja4.write(0,3,L9[i+2])
			hoja4.write(0,4,L9[i+3])
		else:
			hoja1.write(i,0,L0[i])
			hoja1.write(i,1,'C'+'0'+str(i))
			hoja1.write(i,2,L1[i])
			hoja1.write(i,3,183)
			hoja1.write(i,4,i)
			hoja1.write(i,5,L1[i]*(5520000-sum(t))/sum(L1))#EAMC
			hoja3.write(i,1,'C'+'0'+str(i))
		c = 1
	c = 0
	for i in range(len(L2)):
		if c == 0:
			hoja2.write(0,0,L6[i])
			hoja2.write(0,1,L6[i+1])
			hoja2.write(0,2,L6[i+2])
			hoja2.write(0,3,L6[i+3])
			hoja2.write(0,4,L6[i+4])
			hoja2.write(0,5,L6[i+5])
			hoja2.write(0,6,L6[i+6])
			hoja2.write(0,7,L6[i+7])
			hoja2.write(0,8,L6[i+8])
			hoja2.write(0,9,L6[i+9])
		else:
			hoja2.write(i,0,L7[i-1])
			hoja2.write(i,1,'V'+'0'+str(i))
			hoja2.write(i,2,'B'+str(i))
			hoja2.write(i,3,L2[i])
			hoja2.write(i,4,0)
			hoja2.write(i,5,precio_venta.precio[i-1]) #precio de venta MC
			hoja2.write(i,9,i)
		c = 1
	archivo.close()
	datos2 = getcwd() + r"\resultados_mecanismo.XLSX" 
	resultado2 = AlgoritmoOptimizador(datos2)
	
print(d)
