"""Create Sample_Sawtooth_DP_Input.sav after installing requirements.

Run:
    python create_sample_sav.py
"""
import pandas as pd
import pyreadstat

rows = {
    "RESPID": [1001,1002,1003,1004,1005,1006,1007,1008,1009,1010],
    "Q1": [1,2,2,1,1,2,1,2,2,1],
    "Q2": [5,4,3,5,2,4,1,5,3,4],
    "Q7_1": [1,0,1,1,0,1,1,0,1,1],
    "Q7_2": [0,1,1,0,1,0,1,1,0,1],
    "Q7_3": [1,1,0,1,0,1,0,1,1,0],
    "Q10_1_1": [5,4,3,5,2,4,1,5,3,4],
    "Q10_1_2": [4,4,2,5,3,4,2,5,3,4],
    "Q10_2_1": [3,5,4,4,2,5,1,4,3,5],
    "Q10_2_2": [4,5,3,4,3,5,2,4,4,5],
    "Q12_OE": ["Good service","Price is high","No comments","Very convenient","More options needed","Good quality","Nothing","Friendly staff","Improve delivery","Satisfied overall"],
    "Q15": [1,2,3,4,5,99,4,3,2,1],
}
df = pd.DataFrame(rows)

column_labels = {
    "RESPID": "Respondent ID",
    "Q1": "Q1. Gender",
    "Q2": "Q2. Overall satisfaction",
    "Q7_1": "Q7. Brands aware - Coca-Cola",
    "Q7_2": "Q7. Brands aware - Pepsi",
    "Q7_3": "Q7. Brands aware - Sprite",
    "Q10_1_1": "Q10. Satisfaction - Service - Quality",
    "Q10_1_2": "Q10. Satisfaction - Service - Speed",
    "Q10_2_1": "Q10. Satisfaction - Product - Quality",
    "Q10_2_2": "Q10. Satisfaction - Product - Availability",
    "Q12_OE": "Q12. Please tell us why",
    "Q15": "Q15. Recommendation",
}

sat = {1:"Very dissatisfied",2:"Dissatisfied",3:"Neither",4:"Satisfied",5:"Very satisfied"}
value_labels = {
    "Q1": {1:"Male",2:"Female"},
    "Q2": sat,
    "Q7_1": {0:"Not selected",1:"Selected"},
    "Q7_2": {0:"Not selected",1:"Selected"},
    "Q7_3": {0:"Not selected",1:"Selected"},
    "Q10_1_1": sat,
    "Q10_1_2": sat,
    "Q10_2_1": sat,
    "Q10_2_2": sat,
    "Q15": {1:"Definitely would",2:"Probably would",3:"Maybe",4:"Probably would not",5:"Definitely would not",99:"DON'T KNOW"},
}

pyreadstat.write_sav(
    df,
    "Sample_Sawtooth_DP_Input.sav",
    file_label="DP Automation synthetic Sawtooth-style test file",
    column_labels=column_labels,
    variable_value_labels=value_labels,
)
print("Created Sample_Sawtooth_DP_Input.sav")
