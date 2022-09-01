import numpy as np
import pandas as pd

ROUND_POINTS = {'ronda-0':0,'ronda-1':2,'ronda-2':1,'ronda-3':1,'ronda-4':1,'ronda-5':1}

def flatten(l):
    return [item for sublist in l for item in sublist]

df = pd.read_excel("data.xlsx")

# Split by participant

participant_splits = df['Participante'].isna()==False
pstart = df.index[participant_splits].tolist()
pend = [x-1 for x in pstart[1:]] + [df.shape[0]]

all_participants=[]
for pa,pb in zip(pstart,pend):
    pdict = {}
    pdf=df.iloc[pa:pb+1]
    pdict['label'] = pdf.iloc[0,0]
    pdict['family'] = pdict['label'][:4]
    round_array = pdf.iloc[:,2:].to_numpy()

    rounds = []
    for round in range(round_array.shape[0]):
        words = [x.replace(' ','_').lower() for x in round_array[round,:].tolist() if isinstance(x,str)]
        if words != []:
            rounds.append(words)

    #flatten list of words
    unique_words = set(flatten(rounds))

    word_points = {}
    for word in unique_words:
        word_points[word]=0
        for i,round in enumerate(rounds):
            if word in round:
                word_points[word]+=ROUND_POINTS['ronda-'+str(i)]

    #pdict['score'] = word_points
    pdict.update(word_points)
    all_participants.append(pdict)

score_df = pd.DataFrame.from_dict(all_participants)
score_df.to_excel('score_df.xlsx')
[print(x) for x in score_df.columns]