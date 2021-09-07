#Rio2016 Results
#Importing libraries and classes
import Country
import xlsxwriter
import matplotlib.pyplot as plt

#Creating the excel with the name and extension
workbook = xlsxwriter.Workbook('Rio2016.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})

#Initialize lists
ranks = []
names = []
golds = []
silvers = []
bronzes = []

#Generating the instances of the country class with darta obtained from the official RIO 2016 site. Only top 10 countries
c1 = Country.Country('1','United States',46,37,38)
c2 = Country.Country('2','Great Britain',27,23,17)
c3 = Country.Country('3','China',26,18,26)
c4 = Country.Country('4','Russia',19,17,20)
c5 = Country.Country('5','Germany',17,10,15)
c6 = Country.Country('6','Japan',12,8,21)
c7 = Country.Country('7','France',10,18,14)
c8 = Country.Country('8','South Korea',9,3,9)
c9 = Country.Country('9','Italy',8,12,8)
c10 = Country.Country('10','Australia',8,11,10)

#Appending rank
ranks.append(c1.rank)
ranks.append(c2.rank)
ranks.append(c3.rank)
ranks.append(c4.rank)
ranks.append(c5.rank)
ranks.append(c6.rank)
ranks.append(c7.rank)
ranks.append(c8.rank)
ranks.append(c9.rank)
ranks.append(c10.rank)

#Appending name
names.append(c1.name)
names.append(c2.name)
names.append(c3.name)
names.append(c4.name)
names.append(c5.name)
names.append(c6.name)
names.append(c7.name)
names.append(c8.name)
names.append(c9.name)
names.append(c10.name)

#Appending gold
golds.append(c1.gold)
golds.append(c2.gold)
golds.append(c3.gold)
golds.append(c4.gold)
golds.append(c5.gold)
golds.append(c6.gold)
golds.append(c7.gold)
golds.append(c8.gold)
golds.append(c9.gold)
golds.append(c10.gold)

#Appending silver
silvers.append(c1.silver)
silvers.append(c2.silver)
silvers.append(c3.silver)
silvers.append(c4.silver)
silvers.append(c5.silver)
silvers.append(c6.silver)
silvers.append(c7.silver)
silvers.append(c8.silver)
silvers.append(c9.silver)
silvers.append(c10.silver)

#Appending bronze
bronzes.append(c1.bronze)
bronzes.append(c2.bronze)
bronzes.append(c3.bronze)
bronzes.append(c4.bronze)
bronzes.append(c5.bronze)
bronzes.append(c6.bronze)
bronzes.append(c7.bronze)
bronzes.append(c8.bronze)
bronzes.append(c9.bronze)
bronzes.append(c10.bronze)

#Function to plot the graphs based on types
def plotgraph(type):
    plt.bar(names,type)
    plt.title('TOP 10 Medals')
    plt.xlabel('Country')
    plt.ylabel('Totals')
    plt.show()

plotgraph(golds)
plotgraph(silvers)
plotgraph(bronzes)

#Data append
data = (
    [c1.rank,c1.name,c1.gold,c1.silver,c1.bronze],
    [c2.rank,c2.name,c2.gold,c2.silver,c2.bronze],
    [c3.rank,c3.name,c3.gold,c3.silver,c3.bronze],
    [c4.rank,c4.name,c4.gold,c4.silver,c4.bronze],
    [c5.rank,c5.name,c5.gold,c5.silver,c5.bronze],
    [c6.rank,c6.name,c6.gold,c6.silver,c6.bronze],
    [c7.rank,c7.name,c7.gold,c7.silver,c7.bronze],
    [c8.rank,c8.name,c8.gold,c8.silver,c8.bronze],
    [c9.rank,c9.name,c9.gold,c9.silver,c9.bronze],
    [c10.rank,c10.name,c10.gold,c10.silver,c10.bronze],
)
#EXCEL rows and columns
row = 0
col = 0

#Filling the columns and rows with the information
for rank, name, gold, silver, bronze in data:

    worksheet.write(0 , 0 ,'Rank',bold)
    worksheet.write(0, 1, 'Country',bold)
    worksheet.write(0, 2, 'Gold',bold)
    worksheet.write(0, 3, 'Silver',bold)
    worksheet.write(0, 4, 'Bronze',bold)
    worksheet.write(row + 1, col, rank)
    worksheet.write(row + 1, col + 1, name)
    worksheet.write(row + 1, col + 2, gold)
    worksheet.write(row + 1, col + 3, silver)
    worksheet.write(row + 1, col + 4, bronze)
    row += 1
    worksheet.write(12, 0, 'Total',bold)
    worksheet.write(12, 2, '=SUM(C2:C12)')
    worksheet.write(12, 3, '=SUM(D2:D12)')
    worksheet.write(12, 4, '=SUM(E2:E12)')

#Close workbook
workbook.close()

