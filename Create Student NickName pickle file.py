#Create Student NickName pickle file
import pickle as cPickle

nicknames = {"SungEun Daniel Yi" : ["Daniel Yi","SungEun Yi"] }

with open('./nicknames.pickle','wb') as file:
    cPickle.dump(nicknames,file,-1)
    file.close()