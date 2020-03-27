ADO_Table_Verify 理论验证 验证成功的算法：
校验1亿个100以内数字的组合，没有发现重复值，随机组合重复率低于 1E-16
def cal_code():
    return [
        cal_one_code([977] , 6 ),
        cal_one_code([29,31]),
        cal_one_code([23,41]),
        cal_one_code([19,47]),
        cal_one_code([17,53]),
        cal_one_code([11,79]),
        cal_one_code([7,137]),
        cal_one_code([5,181]),
        cal_one_code([3,293]),
        cal_one_code([2,463]),
        cal_one_code([5,11,17]),
        cal_one_code([3,13,23]),
    ]

def cal_one_code(factors, len_code = 2):
    code_sum = 0
    max_code = 10 ** len_code
    for x in list_x:
        if x > 0 :
            code_to_add = 1
            for factor in factors:
                code_to_add = code_to_add * (x % factor + 1)
            code_sum = code_sum + code_to_add
    return code_sum % max_code


select 
    format(sum(ID mod 977) mod 1000000 ,"000000") & "-" &
    format(sum((ID mod 29) * (ID mod 31)) mod 100 ,"00") & "-" &
    format(sum((ID mod 23) * (ID mod 41)) mod 100 ,"00") & "-" &
    format(sum((ID mod 19) * (ID mod 47)) mod 100 ,"00") & "-" &
    format(sum((ID mod 17) * (ID mod 53)) mod 100 ,"00") & "-" &
    format(sum((ID mod 11) * (ID mod 79)) mod 100 ,"00") & "-" &
    format(sum((ID mod 7) * (ID mod 137)) mod 100 ,"00") & "-" &
    format(sum((ID mod 5) * (ID mod 181)) mod 100 ,"00") & "-" &
    format(sum((ID mod 3) * (ID mod 293)) mod 100 ,"00") & "-" &
    format(sum((ID mod 2) * (ID mod 463)) mod 100 ,"00") & "-" &
    format(sum((ID mod 5) * (ID mod 11) * (ID mod 17)) mod 100 ,"00") & "-" &
    format(sum((ID mod 3) * (ID mod 13) * (ID mod 23)) mod 100 ,"00") as Check_Code
from ;


select sum(ID mod 977) mod 1000000 as Check_Code , int(ID / 983) mod 2 as Check_Group
from T1
order by Check_Group;