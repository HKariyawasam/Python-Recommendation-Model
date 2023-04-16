from flask import Flask,jsonify,request
from src.models.physical_activity_models.collabortive_based_model import collabUser
from src.models.physical_activity_models.content_based_model import contentClustering
from src.data.dataset import updateUserCompletionRates
app = Flask(__name__)



@app.route('/')
def index():
    return jsonify({'name':'Hasani',
                    'status':'Server Working'})


@app.route('/api/v1/activities/users/<int:userId>',methods=['GET'])
def get_collaborative_preference(userId):  
    list = collabUser(userId)
    if(len(list) !=0):
        return jsonify(list)
    else:
        return jsonify({'name':"An Error Occured",
                    'status':404})

                    

@app.route('/api/v1/activities/<int:activityIndex>',methods=['GET'])
def get_clustered_physical_activities(activityIndex):  
    list = contentClustering(activityIndex)
    if(len(list) !=0):
        return jsonify(list)
    else:
        return jsonify({'name':"An Error Occured",
                    'status':404})
    
@app.route('/api/v1/activities/completerates',methods=['POST'])
def update_completerate_sheet():  
    updateUserCompletionRates(request.json)
    return jsonify({'name':"Success",
                    'status':200})



