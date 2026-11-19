import os
ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA=os.path.join(ROOT,'data'); WORK=os.path.join(ROOT,'work'); OUT=os.path.join(ROOT,'salidas')
SRC=os.path.join(DATA,'CALENDARIO2027REVISAR.xlsx')
for p in (WORK,OUT): os.makedirs(p,exist_ok=True)
