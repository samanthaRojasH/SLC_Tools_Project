import pandas as pd
import glob

path = "C:/Ruta/donde/estan/Documentos/aUnir/ParaPonerResultadodeUnir/AsisAuto/"
files = path + "*.csv"
output = pd.DataFrame()

for file in glob.glob(files):
    data_temporal = pd.read_csv(file)
    output = output.append(data_temporal,ignore_index=True)

    output["dummy"]=1    
    output_2 = pd.pivot_table(output,index="Participantes",values="dummy",aggfunc="sum")
    output_2.to_csv(path + "asistencia.csv")
