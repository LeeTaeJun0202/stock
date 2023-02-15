import pandas as pd

date = pd.read_excel("납기_2.xlsx", engine="openpyxl")
price = pd.read_excel("가격_2.xlsx", engine="openpyxl")
demand = pd.read_excel("수요_2.xlsx", engine="openpyxl")
holding_cost = pd.read_excel("유지비용.xlsx", engine="openpyxl")
holding_cost=holding_cost.iloc[:,1:2]
order = pd.read_excel("주문날짜.xlsx", engine="openpyxl")

m = 12
count_A=[45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45,45]
count_A=pd.DataFrame(count_A)
count_A.columns=["입고 수량"]
order_a=order.iloc[0:m,0:1]+pd.DateOffset(weeks=date.iloc[0,0])
order_aa=order.iloc[m:2*m,0:1]+pd.DateOffset(weeks=date.iloc[1,0])
order_aaa=order.iloc[2*m:3*m,0:1]+pd.DateOffset(weeks=date.iloc[2,0])
order_A=pd.concat([order_a,order_aa,order_aaa], axis=0)
order_A=order_A.sort_values(by='날짜')
order_A=pd.concat([order_A,count_A], axis=1)

count_B=[35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35]
count_B=pd.DataFrame(count_B)
count_B.columns=["입고 수량"]
order_b=order.iloc[0:m,0:1]+pd.DateOffset(weeks=date.iloc[0,1])
order_bb=order.iloc[m:2*m,0:1]+pd.DateOffset(weeks=date.iloc[1,1])
order_bbb=order.iloc[2*m:3*m,0:1]+pd.DateOffset(weeks=date.iloc[2,1])
order_B=pd.concat([order_b,order_bb,order_bbb], axis=0)
order_B=order_B.sort_values(by='날짜')
order_B=pd.concat([order_B,count_B], axis=1)

count_C=[55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55,55]
count_C=pd.DataFrame(count_C)
count_C.columns=["입고 수량"]
order_c=order.iloc[0:m,0:1]+pd.DateOffset(weeks=date.iloc[0,2])
order_cc=order.iloc[m:2*m,0:1]+pd.DateOffset(weeks=date.iloc[1,2])
order_ccc=order.iloc[2*m:3*m,0:1]+pd.DateOffset(weeks=date.iloc[2,2])
order_C=pd.concat([order_c,order_cc,order_ccc], axis=0)
order_C=order_C.sort_values(by='날짜')
order_C=pd.concat([order_C,count_C], axis=1)

on_hand_A = 230
stock_A = demand.iloc[:,0:2]
stock_A['입고 수량(A)'] = 0
for i,l in zip(order_A.iloc[:,0],order_A.index):

    for j,k in zip(demand.iloc[:,0],demand.index):

        if j>i:

            stock_A.loc[k,'입고 수량(A)']+=order_A.iloc[l,1]
            break

stock_A['재고(A)'] = 0

stock_A.loc[0,'재고(A)']=on_hand_A-stock_A.loc[0,'수요(A)']

for i,j in zip(stock_A.loc[:,'재고(A)'],stock_A.index):
    if j==35:
        break
    stock_A.loc[j+1,'재고(A)']=i-stock_A.loc[j+1,'수요(A)']+stock_A.loc[j+1,'입고 수량(A)']

stock_A['재고량 총합(A)'] = 0

stock_A.loc[0,'재고량 총합(A)']=stock_A.loc[0,'재고(A)']

for j in stock_A.index:
    if j==35:
        break
    stock_A.loc[j+1,'재고량 총합(A)']=stock_A.loc[j,'재고량 총합(A)']+stock_A.loc[j+1,'재고(A)']

on_hand_B = 600
stock_B = demand[['날짜','수요(B)']]
stock_B['입고 수량(B)'] = 0
for i,l in zip(order_B.iloc[:,0],order_B.index):

    for j,k in zip(demand.iloc[:,0],demand.index):

        if j>i:

            stock_B.loc[k,'입고 수량(B)']+=order_B.iloc[l,1]
            break

stock_B['재고(B)'] = 0

stock_B.loc[0,'재고(B)']=on_hand_B-stock_B.loc[0,'수요(B)']

for i,j in zip(stock_B.loc[:,'재고(B)'],stock_B.index):
    if j==35:
        break
    stock_B.loc[j+1,'재고(B)']=i-stock_B.loc[j+1,'수요(B)']+stock_B.loc[j+1,'입고 수량(B)']

stock_B['재고량 총합(B)'] = 0

stock_B.loc[0,'재고량 총합(B)']=stock_B.loc[0,'재고(B)']

for j in stock_B.index:
    if j==35:
        break
    stock_B.loc[j+1,'재고량 총합(B)']=stock_B.loc[j,'재고량 총합(B)']+stock_B.loc[j+1,'재고(B)']

on_hand_C = 200
stock_C = demand[['날짜','수요(C)']]
stock_C['입고 수량(C)'] = 0

for i,l in zip(order_C.iloc[:,0],order_C.index):

    for j,k in zip(demand.iloc[:,0],demand.index):

        if j>i:

            stock_C.loc[k,'입고 수량(C)']+=order_C.iloc[l,1]
            break

stock_C['재고(C)'] = 0

stock_C.loc[0,'재고(C)']=on_hand_C-stock_C.loc[0,'수요(C)']

for i,j in zip(stock_C.loc[:,'재고(C)'],stock_C.index):
    if j==35:
        break
    stock_C.loc[j+1,'재고(C)']=i-stock_C.loc[j+1,'수요(C)']+stock_C.loc[j+1,'입고 수량(C)']

stock_C['재고량 총합(C)'] = 0

stock_C.loc[0,'재고량 총합(C)']=stock_C.loc[0,'재고(C)']

for j in stock_C.index:
    if j==35:
        break
    stock_C.loc[j+1,'재고량 총합(C)']=stock_C.loc[j,'재고량 총합(C)']+stock_C.loc[j+1,'재고(C)']

holding_A=holding_cost
holding_A['유지비용*매달 재고량 총합']=0
for i,j in zip(holding_A.iloc[:,0],holding_A.index):
    holding_A.loc[j,'유지비용*매달 재고량 총합']=stock_A.iloc[-1,-1]*i
holding_A['목적함수(A)']=0
def list_chunk(lst, n):
    return [lst[i:i+n] for i in range(0, len(lst), n)]
chunk_A = list_chunk(count_A, 12) #1년이니까 한 리스트에 12개의 값이 들어가게끔.

a=[]
sum_A=0
for i in price.index:
    sum_A=0
    for j in chunk_A[i].iloc[:,0]:
        sum_A+=j
    a.append(sum_A)
total_A=0
for i in price.index:
    total_A+=price.iloc[i,0]*a[i]
for i,j in zip(holding_A.iloc[:,1],holding_A.index):
    holding_A.loc[j,'목적함수(A)']=i+total_A

holding_B=holding_cost
holding_B['유지비용*매달 재고량 총합']=0
for i,j in zip(holding_B.iloc[:,0],holding_B.index):
    holding_B.loc[j,'유지비용*매달 재고량 총합']=stock_B.iloc[-1,-1]*i
holding_B['목적함수(B)']=0

chunk_B = list_chunk(count_B, 12) #1년이니까 한 리스트에 12개의 값이 들어가게끔.
b=[]
sum_B=0
for i in price.index:
    sum_B=0
    for j in chunk_B[i].iloc[:,0]:
        sum_B+=j
    b.append(sum_B)
total_B=0
for i in price.index:
    total_B+=price.iloc[i,1]*b[i]
for i,j in zip(holding_B.iloc[:,1],holding_B.index):
    holding_B.loc[j,'목적함수(B)']=i+total_B

holding_C=holding_cost
holding_C['유지비용*매달 재고량 총합']=0
for i,j in zip(holding_C.iloc[:,0],holding_C.index):
    holding_C.loc[j,'유지비용*매달 재고량 총합']=stock_C.iloc[-1,-1]*i
holding_C['목적함수(C)']=0

chunk_C = list_chunk(count_C, 12) #1년이니까 한 리스트에 12개의 값이 들어가게끔.
c=[]
sum_C=0
for i in price.index:
    sum_C=0
    for j in chunk_C[i].iloc[:,0]:
        sum_C+=j
    c.append(sum_C)
total_C=0
for i in price.index:
    total_C+=price.iloc[i,2]*c[i]
for i,j in zip(holding_C.iloc[:,1],holding_C.index):
    holding_C.loc[j,'목적함수(C)']=i+total_C

target=holding_C.iloc[:,2:5]
print(target)