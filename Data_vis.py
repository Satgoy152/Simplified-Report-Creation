# %%
import pandas as pd
import requests
from IPython.display import HTML
def get_file_details(file_name):
    #change file name

    #get file information
    file_details = file_name.split("-")
    value_type = file_details[1]
    organization_type = file_details[3]

    df = pd.read_excel(file_name + ".xlsx")
    data_frames = []
    return df, value_type, organization_type

# %%
def new_str(str):
    activity = str.split(",")
    new = ""
    for act in activity:
        new = new + act + "\015"
    return new

# %%
def get_data(df, value_type, organization_type):
    data_frames = []

    for r in range(len(df["Date"].values.tolist())):
        
        if(value_type == 'quantity'):
            prev_quant = [df.loc[r,'Order_Quantity in Previous Week']]
            quant = [df.loc[r,'Order_Quantity in This Week']]
            quant_diff = [df.loc[r,'Difference in the Order_Quantity (Previous vs This)']]
            gain = [df.loc[r,'Gain in Order_Quantity (Previous vs This)']]
            loss = [df.loc[r,'Loss in Order_Quantity (Previous vs This)']]
            top_gain = [df.loc[r,'Total Gain from TopN Gainers']]
            top_loss = [df.loc[r,'Total Loss from TopN Losers']]
            perc_rise = [str(df.loc[r,'Percentage of Rise from TopN Gainers']) + "%"]
            perc_drop = [str(df.loc[r,'Percentage of Drop from TopN Losers']) + "%"]
        else:
            prev_rev = [df.loc[r,'Revenue in Previous Week']]
            rev = [df.loc[r,'Revenue in This Week']]
            rev_diff = [df.loc[r,'Difference in the Revenue (Previous vs This)']]
            gain = [df.loc[r,'Gain in Revenue (Previous vs This)']]
            loss = [df.loc[r,'Loss in Revenue (Previous vs This)']]
            top_gain = [df.loc[r,'Total Gain from TopN Gainers']]
            top_loss = [df.loc[r,'Total Loss from TopN Losers']]
            perc_rise = [str(df.loc[r,'Percentage of Rise from TopN Gainers']) + "%"]
            perc_drop = [str(df.loc[r,'Percentage of Drop from TopN Losers']) + "%"]
        
        #defining arrays
        date = [df.loc[r,'Date']]
        losers = []
        loser_urls = []
        loser_walmarturls = []
        loser_ids = []
        loser_metrics = []
        loser_trends = []
        loser_downs = []
        loser_ups = []
        loser_act = []

        gainers = []
        gainer_urls = []
        gainer_walmarturls = []
        gainer_ids = []
        gainer_metrics = []
        gainer_trends = []
        gainer_downs = []
        gainer_ups = []
        gainer_act = []
        #grabbing data from excel file
        if value_type == 'revenue':
            for i in range(10):
                losers.append(df.loc[r,'Top Losers #' + str(i+1) + ' Product Type'])
                loser_ids.append(str(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID']))
                loser_walmarturls.append("https://www.walmart.com/ip/" + str(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID']))
                loser_urls.append("https://app.ezdia.com/catalog?query=(%20item_id%20EQ%20" + str(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID']) +"%20)")
                loser_metrics.append(df.loc[r,'Top Losers #' + str(i+1) + ' Difference in Revenue'])
                loser_trends.append(df.loc[r,'Top Losers #' + str(i+1) + ' Recent Revenue Trend'])
                loser_downs.append(df.loc[r,'Top Losers #' + str(i+1) + ' Number of Downs'])
                loser_ups.append(df.loc[r,'Top Losers #' + str(i+1) + ' Number of UPs'])
                loser_act.append(new_str(str(df.loc[r,'Top Losers #' + str(i+1) + ' Recent Change Events'])))

                gainers.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Product Type'])
                gainer_ids.append(str(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID']))
                gainer_walmarturls.append("https://www.walmart.com/ip/" + str(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID']))
                gainer_urls.append("https://app.ezdia.com/catalog?query=(%20item_id%20EQ%20" + str(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID']) +"%20)")
                gainer_metrics.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Difference in Revenue'])
                gainer_trends.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Recent Revenue Trend'])
                gainer_downs.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Number of Downs'])
                gainer_ups.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Number of UPs'])
                gainer_act.append(new_str(str(df.loc[r,'Top Gainers #' + str(i+1) + ' Recent Change Events'])))
    
            #organizing data into a dataframe
            dates_dic = {'Date': date}
            df2 = pd.DataFrame(data = dates_dic)
            df2['Date'] = pd.to_datetime(df2['Date']).dt.date
            df2 = df2.T
            df2 = df2.reset_index()

            rev_loss = {'Revenue in Previous Week': prev_rev, 'Revenue in This Week': rev, 
                        'Difference in the Revenue (Previous vs This)': rev_diff, 'Gain in Revenue (Previous vs This)': gain, 
                        'Loss in Revenue (Previous vs This)': loss, 'Total Gain from TopN Gainers': top_gain, 
                        'Total Loss from TopN Losers': top_loss, 'Percentage of Rise from TopN Gainers': perc_rise, 
                        'Percentage of Drop from TopN Losers': perc_drop}
            df3 = pd.DataFrame(data = rev_loss)
            df3 = df3.T
            df3 = df3.reset_index()
            
            losers_data = {'Item ID': loser_ids, 'Product Type': losers, 'Difference in Revenue': loser_metrics, 
                            'Recent Revenue Trend': loser_trends, 'Number of Downs': loser_downs, 'Number of Ups': loser_ups,
                            'Change Events': loser_act, 'Walmart URLs': loser_walmarturls, 'Crewmachine URLs': loser_urls}
            
            df4 = pd.DataFrame(data=losers_data)
            HTML(df4.to_html(render_links=True, escape=False))

            gainers_data = {'Item ID': gainer_ids, 'Product Type': gainers, 'Difference in Revenue': gainer_metrics,
                            'Recent Revenue Trend': gainer_trends, 'Number of Downs': gainer_downs, 'Number of Ups': gainer_ups,
                            'Change Events': gainer_act, 'Walmart URLs': gainer_walmarturls, 'Crewmachine URLs': gainer_urls}
            df5 = pd.DataFrame(data=gainers_data)
            HTML(df5.to_html(render_links=True, escape=False))

        else:
            for i in range(10):
                losers.append(df.loc[r,'Top Losers #' + str(i+1) + ' Product Type'])
                loser_ids.append(str(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID']))
                loser_walmarturls.append("https://www.walmart.com/ip/" + str(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID']))
                loser_urls.append("https://app.ezdia.com/catalog?query=(%20item_id%20EQ%20" + str(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID']) +"%20)")
                loser_metrics.append(df.loc[r,'Top Losers #' + str(i+1) + ' Difference in Order_Quantity'])
                loser_trends.append(df.loc[r,'Top Losers #' + str(i+1) + ' Recent Order_Quantity Trend'])
                loser_downs.append(df.loc[r,'Top Losers #' + str(i+1) + ' Number of Downs'])
                loser_ups.append(df.loc[r,'Top Losers #' + str(i+1) + ' Number of UPs'])
                loser_act.append(new_str(str(df.loc[r,'Top Losers #' + str(i+1) + ' Recent Change Events'])))

                gainers.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Product Type'])
                gainer_ids.append(str(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID']))
                gainer_walmarturls.append("https://www.walmart.com/ip/" + str(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID']))
                gainer_urls.append("https://app.ezdia.com/catalog?query=(%20item_id%20EQ%20" + str(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID']) +"%20)")
                gainer_metrics.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Difference in Order_Quantity'])
                gainer_trends.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Recent Order_Quantity Trend'])
                gainer_downs.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Number of Downs'])
                gainer_ups.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Number of UPs'])
                gainer_act.append(new_str(str(df.loc[r,'Top Gainers #' + str(i+1) + ' Recent Change Events'])))

            #organizing data into a dataframe
            dates_dic = {'Date': date}
            df2 = pd.DataFrame(data = dates_dic)
            df2['Date'] = pd.to_datetime(df2['Date']).dt.date
            df2 = df2.T
            df2 = df2.reset_index()
            

            quant_loss = {'Quantity in Previous Week': prev_quant, 'Quantity in This Week': quant, 
                                'Difference in the Order_Quantity (Previous vs This)': quant_diff, 
                                'Gain in Order_Quantity (Previous vs This)': gain, 'Loss in Order_Quantity (Previous vs This)': loss, 
                                'Total Gain from TopN Gainers': top_gain, 'Total Loss from TopN Losers': top_loss,
                                'Percentage of Rise from TopN Gainers': perc_rise, 'Percentage of Drop from TopN Losers': perc_drop}
            df3 = pd.DataFrame(data = quant_loss)
            df3 = df3.T
            df3 = df3.reset_index()


            losers_data = {'Item ID': loser_ids, 'Product Type': losers, 'Difference in Units Sold': loser_metrics, 
                            'Recent Units Sold Trend': loser_trends, 'Number of Downs': loser_downs, 'Number of Ups': loser_ups, 
                            'Change Events': loser_act, 'Walmart URLs': loser_walmarturls, 'Crewmachine URLs': loser_urls}
            df4 = pd.DataFrame(data=losers_data)
            HTML(df4.to_html(render_links=True, escape=False))
            df4 = df4.style.set_caption('Top Losers')

            gainers_data = {'Item ID': gainer_ids, 'Product Type': gainers, 'Difference in Units Sold': gainer_metrics, 
                            'Recent Units Sold Trend': gainer_trends, 'Number of Downs': gainer_downs, 'Number of Ups': gainer_ups, 
                            'Change Events': gainer_act, 'Walmart URLs': gainer_walmarturls, 'Crewmachine URLs': gainer_urls}
            df5 = pd.DataFrame(data=gainers_data)
            HTML(df5.to_html(render_links=True, escape=False))


        #add all data frames to a list
        data_frames.append(df2)
        data_frames.append(df3)
        data_frames.append(df4)
        data_frames.append(df5)
    return value_type, data_frames

# %%
from email import header
from textwrap import wrap
from matplotlib.pyplot import title

from pyparsing import conditionAsParseAction

writer = pd.ExcelWriter("wcp280-simplified.xlsx", engine='xlsxwriter')

def excel_writer(value_type, data_frames, s):
    

    data_frames[0].to_excel(writer, sheet_name='Sheet' + str(s), startrow= 0, index=False, startcol=0, header=False)

    wb = writer.book
    ws = writer.sheets['Sheet' + str(s)]

    l = 0
    data_frames.reverse()
    for i in range(0, len(data_frames), 4):

        data_frames[i+3].to_excel(writer, sheet_name='Sheet' + str(s), startrow= l, index=False, startcol=0, header=False)
        data_frames[i+2].to_excel(writer, sheet_name='Sheet' + str(s), startrow= l + 2, index=False, startcol=0, header=False)
        
        data_frames[i+1].to_excel(writer, sheet_name='Sheet' + str(s), startrow= l + 16, index=False, startcol=0)
        ws.write_string(l+15,0, 'Top Losers')
        #ws.insert_image('H'+str(l+16), 'walmart_logo.png', {'x_scale': 0.1, 'y_scale': 0.1})
        #ws.insert_image('I'+str(l+16), 'optiwise_logo.png', {'x_scale': 0.1, 'y_scale': 0.1})
        data_frames[i].to_excel(writer, sheet_name='Sheet' + str(s), startrow= l + 16, index=False, startcol=11)
        ws.write_string(l+15, 11, 'Top Gainers')
        #ws.insert_image('S'+str(l+16), 'walmart_logo.png')
        #ws.insert_image('T'+str(l+16), 'optiwise_logo.png')
        i += 4
        l += 32

    #formatting
    wrap_format = wb.add_format({'text_wrap': True})
    money_fmt = wb.add_format({'num_format': '$#,###.##'})
    ws.set_column('A:A', 30, wrap_format)
    ws.set_column('B:T', 20)
    if value_type == 'revenue':
        ws.set_column('B:D', 20, money_fmt)
        ws.set_column('N:N', 20, money_fmt)
        ws.set_column('E:M', 20)
        ws.set_column('O:U', 20)

# %%
#running the script
#change file names to the file you want to run
file_list = ['wcp280-quantity-N10-item', 'wcp280-revenue-N10-item', 'wcp280-quantity-N10-product', 'wcp280-revenue-N10-product']

#leave rest as is if no changes are need to be made
s = 1
for file in file_list:
    df, value_type, organization_type = get_file_details(file)
    value_type, data_frames = get_data(df, value_type, organization_type)
    excel_writer(value_type, data_frames, s)
    print(value_type, organization_type)
    s+=1
writer.save()


