import numpy as np
import pandas as pd
import os
import unicodedata

output = 'results'
os.makedirs(output,exist_ok=True)

def remove_accents(input_str):
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

ROUND_POINTS = {'ronda-0':0,'ronda-1':2,'ronda-2':1,'ronda-3':1,'ronda-4':1,'ronda-5':1}

def flatten(l):
    return [item for sublist in l for item in sublist]

df = pd.read_excel("data.xlsx")

# Split by participant

participant_splits = df['Participante'].isna()==False
pstart = df.index[participant_splits].tolist()
pend = [x-1 for x in pstart[1:]] + [df.shape[0]]

all_words = [] # as originally written
all_participants=[]
for pa,pb in zip(pstart,pend):
    pdict = {}
    pdf=df.iloc[pa:pb+1]
    pdict['label'] = pdf.iloc[0,0]
    pdict['family'] = pdict['label'][:4]
    round_array = pdf.iloc[:,2:].to_numpy()

    rounds = []
    for round in range(round_array.shape[0]):
        words = [x for x in round_array[round,:].tolist() if isinstance(x,str)]
        if words != []:
            rounds.append([remove_accents(x.replace('ñ','%n%').replace(' ','_').lower()).replace('%n%','ñ') for x in words])
            all_words += words

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
score_df.to_excel(os.path.join(output,'WordScoresPerParticipant.xlsx'))

all_words = list(set(all_words))
all_words.sort()

all_words = pd.Series(all_words).to_frame()
all_words.to_html(os.path.join(output,'OriginalWords.html'),encoding="cp1252")
all_words.to_excel(os.path.join(output,'OriginalWords.xlsx'))

word_sum = score_df.iloc[:,2:].sum().sort_values(ascending=False).to_frame()
word_sum.to_html(os.path.join(output,'WordSum-AllFamilies.html'),encoding="cp1252")
word_sum.to_excel(os.path.join(output,'WordSum-AllFamilies.xlsx'))
#word_sum.to_markdown('WordSum.md')

# by family

fam_scores = score_df.groupby('family')

for gr in fam_scores:
    family = gr[0]
    dft = gr[1]
    print(dft)
    word_sum = dft.iloc[:,2:].sum().sort_values(ascending=False).to_frame()
    #word_sum.to_markdown(f'WordSum-{family}.md')
    word_sum.to_html(os.path.join(output,f'WordSum-{family}.html'),encoding="cp1252")
    word_sum.to_excel(os.path.join(output,f'WordSum-{family}.xlsx'))