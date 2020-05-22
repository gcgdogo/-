#coding:UTF-8
from __future__ import division, print_function, absolute_import, unicode_literals

from random import randint
Pattern_Library = {'all':[]}    # all 作为全部列表   
Color_Library = [
    [0,0,0],
    [0.5,0,0],
    [0,0.5,0],
    [0,0,0.5],
]

Current_Value = 0

def Rand_Color():
    return Color_Library[
        randint(0 , len(Color_Library) - 1)
    ]

def Fill_Rand_Color(PatternData):
    Pattern_Nodes = []
    for NodeX in PatternData['nodes']:
        NodeX = NodeX.copy()
        if 'color' in NodeX:
            if NodeX['color'] == []:
                NodeX['color'] = Rand_Color()
        Pattern_Nodes.append(NodeX)
    return {
        'value' : PatternData['value'] ,
        'nodes' : Pattern_Nodes
    }

def Rand_Pattern(target_value = 'all'):
    #选区随机数据
    PatternData = Pattern_Library[target_value][
        randint( 0 , len(Pattern_Library[target_value]) - 1 )
    ]
    #填写随机颜色
    PatternData = Fill_Rand_Color(PatternData)
    
    return PatternData

def Add_Data(*args):
    Pattern_Nodes = []
    for NodeX in args :
        if not('color') in NodeX :  #如果第一项没有color  添加color:[]
            if len(Pattern_Nodes) == 0 :
                NodeX['color'] = []

        Pattern_Nodes.append(NodeX)
    
    if not Current_Value in Pattern_Library:
        Pattern_Library[Current_Value] = []
    
    PatternData = {
        'value':Current_Value ,
        'nodes':Pattern_Nodes 
    }

    Pattern_Library[Current_Value].append(PatternData)
    Pattern_Library['all'].append(PatternData)



######################################################
Current_Value = 1
######################################################

Add_Data({"x":[0] , 'y':[0]})
Add_Data({"x":[-0.5] , 'y':[0.5]})
Add_Data({"x":[0.5] , 'y':[-0.5]})




######################################################
Current_Value = 2
######################################################

Add_Data({"x":[-0.3,0.3] , 'y':[0]})
Add_Data({"x":[0] , 'y':[-0.3,0.3]})

Add_Data({"x":[-0.2,0.2] , 'y':[-0.2,0.2]})
Add_Data({"x":[0.2,-0.2] , 'y':[-0.2,0.2]})




######################################################
Current_Value = 3
######################################################

Add_Data({"x":[-0.5,0,0.5] , 'y':[0]})
Add_Data({"x":[0] , 'y':[-0.5,0,0.5]})

Add_Data({"x":[-0.5,0,0.5] , 'y':[-0.5,0,0.5]})
Add_Data({"x":[0.5,0,-0.5] , 'y':[-0.5,0,0.5]})

Add_Data({'r':[0.7],'t':[0,1/3,2/3]})
Add_Data({'r':[0.7],'t':[1/6,1/2,5/6]})





######################################################
######################################################
######################################################

if __name__ =='__main__':
    print(Rand_Pattern())
    for key , item_list in Pattern_Library.items():
        print('-----------------------------------')
        print('key : {}'.format(key))

        for item_node in item_list:
            print("   {}".format(item_node))