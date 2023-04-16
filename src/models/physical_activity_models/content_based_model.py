import pandas as pd
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.metrics import pairwise_distances
import matplotlib.pyplot as plt
from sentence_transformers import SentenceTransformer
import seaborn as sns
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from src.data.dataset import readActivityDataFrame

plt.style.use('ggplot')
plt.rcParams['font.family'] = 'sans-serif' 
plt.rcParams['font.serif'] = 'Ubuntu' 
plt.rcParams['font.monospace'] = 'Ubuntu Mono' 
plt.rcParams['font.size'] = 14 
plt.rcParams['axes.labelsize'] = 12 
plt.rcParams['axes.labelweight'] = 'bold' 
plt.rcParams['axes.titlesize'] = 12 
plt.rcParams['xtick.labelsize'] = 12 
plt.rcParams['ytick.labelsize'] = 12 
plt.rcParams['legend.fontsize'] = 12 
plt.rcParams['figure.titlesize'] = 12 
plt.rcParams['image.cmap'] = 'jet' 
plt.rcParams['image.interpolation'] = 'none' 
plt.rcParams['figure.figsize'] = (12, 10) 
plt.rcParams['axes.grid']=True
plt.rcParams['lines.linewidth'] = 2 
plt.rcParams['lines.markersize'] = 8
colors = ['xkcd:pale orange', 'xkcd:sea blue', 'xkcd:pale red', 'xkcd:sage green', 'xkcd:terra cotta', 'xkcd:dull purple', 'xkcd:teal', 'xkcd: goldenrod', 'xkcd:cadet blue',
'xkcd:scarlet']

def contentClustering(activityIndex):
    data = readActivityDataFrame()
    X = np.array(data.types)

    data = data[['types','category','title']]

    text_data = X
    model = SentenceTransformer('distilbert-base-nli-mean-tokens')
    embeddings = model.encode(text_data, show_progress_bar=True)

    X = np.array(embeddings)
    n_comp = 5
    pca = PCA(n_components=n_comp)
    pca.fit(X)
    pca_data = pd.DataFrame(pca.transform(X))


    cos_sim_data = pd.DataFrame(cosine_similarity(X))
    def give_recommendations(index,print_recommendation = False,print_recommendation_plots= False,print_category =False):
        index_recomm =cos_sim_data.loc[index].sort_values(ascending=False).index.tolist()[1:6]
        activities_recomm =  data['title'].loc[index_recomm].values
        results = index_recomm
        result = {'Activities':activities_recomm,'Index':index_recomm}
        if print_recommendation==True:
            print('The attempted activity is : %s \n'%(data['title'].loc[index]))
            k=1
            for activity in activities_recomm:
                print('The number %i recommended activity is e: %s \n'%(k,activity))
        if print_recommendation_plots==True:
            print('The plot of the attempted activity is :\n %s \n'%(data['types'].loc[index]))
            k=1
            for q in range(len(activities_recomm)):
                plot_q = data['types'].loc[index_recomm[q]]
                print('The plot of the number %i recommended activity is :\n %s \n'%(k,plot_q))
            k=k+1
        if print_category==True:
            print('The category of the attempted activity is :\n %s \n'%(data['category'].loc[index]))
            k=1
            for q in range(len(activities_recomm)):
                plot_q = data['category'].loc[index_recomm[q]]
                print('The plot of the number %i recommended activity is :\n %s \n'%(k,plot_q))
            k=k+1
        return results

    return give_recommendations(activityIndex,True)
