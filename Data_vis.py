# %%
import pandas as pd
import requests
from IPython.display import HTML

#change file name
file_name = "wcp280-revenue-N10-item"

#get file information
file_details = file_name.split("-")
value_type = file_details[1]
organization_type = file_details[3]

df = pd.read_excel(file_name + ".xlsx")
data_frames = []
dates = df["Date"].values.tolist()

# %%
def new_str(str):
    activity = str.split(",")
    new = ""
    for act in activity:
        new = new + act + "\015"
    return new

# %%
def get_data(df, value_type, organization_type):
    for r in range(len(df["Date"].values.tolist())):
        
        if(value_type == 'quantity'):
            gain = [df.loc[r,'Gain in Order_Quantity (Previous vs This)']]
            loss = [df.loc[r,'Loss in Order_Quantity (Previous vs This)']]
        else:
            gain = [df.loc[r,'Gain in Revenue (Previous vs This)']]
            loss = [df.loc[r,'Loss in Revenue (Previous vs This)']]
        
        
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
        if value_type == 'revenue':
            for i in range(5):
                losers.append(df.loc[r,'Top Losers #' + str(i+1) + ' Product Type'])
                loser_ids.append(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID'])
                loser_walmarturls.append("https://www.walmart.com/ip/" + str(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID']))
                loser_urls.append("https://app.ezdia.com/catalog?query=(%20item_id%20EQ%20" + str(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID']) +"%20)")
                loser_metrics.append(df.loc[r,'Top Losers #' + str(i+1) + ' Difference in Revenue'])
                loser_trends.append(df.loc[r,'Top Losers #' + str(i+1) + ' Recent Revenue Trend'])
                loser_act.append(new_str(str(df.loc[r,'Top Losers #' + str(i+1) + ' Recent Change Events'])))

                gainers.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Product Type'])
                gainer_ids.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID'])
                gainer_walmarturls.append("https://www.walmart.com/ip/" + str(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID']))
                gainer_urls.append("https://app.ezdia.com/catalog?query=(%20item_id%20EQ%20" + str(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID']) +"%20)")
                gainer_metrics.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Difference in Revenue'])
                gainer_trends.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Recent Revenue Trend'])
                gainer_act.append(new_str(str(df.loc[r,'Top Gainers #' + str(i+1) + ' Recent Change Events'])))

            losers_data = {'Losers': losers, 'IDs': loser_ids, 'Walmart URLs': loser_walmarturls, 'Crewmachine URLs': loser_urls, 'Revenues': loser_metrics, 'Trends': loser_trends, 'Activity': loser_act}
            df2 = pd.DataFrame(data=losers_data)
            HTML(df2.to_html(render_links=True, escape=False))


            gainers_data = {'Gainers': gainers, 'IDs': gainer_ids, 'Walmart URLs': gainer_walmarturls, 'Crewmachine URLs': gainer_urls, 'Revenues': gainer_metrics, 'Trends': gainer_trends, 'Activity': gainer_act}
            df3 = pd.DataFrame(data=gainers_data)
            HTML(df3.to_html(render_links=True, escape=False))


            dates_and_rev_loss = {'Date': date, 'Loss': loss}
            df4 = pd.DataFrame(data = dates_and_rev_loss)

            dates_and_rev_gain = {'Date': date, 'Gain': gain}
            df5 = pd.DataFrame(data = dates_and_rev_gain)
        else:
            for i in range(5):
                losers.append(df.loc[r,'Top Losers #' + str(i+1) + ' Product Type'])
                loser_ids.append(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID'])
                loser_walmarturls.append("https://www.walmart.com/ip/" + str(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID']))
                loser_urls.append("https://app.ezdia.com/catalog?query=(%20item_id%20EQ%20" + str(df.loc[r,'Top Losers #' + str(i+1) + ' Item ID']) +"%20)")
                loser_metrics.append(df.loc[r,'Top Losers #' + str(i+1) + ' Difference in Order_Quantity'])
                loser_trends.append(df.loc[r,'Top Losers #' + str(i+1) + ' Recent Order_Quantity Trend'])
                loser_downs.append(df.loc[r,'Top Losers #' + str(i+1) + ' Number of Downs'])
                loser_ups.append(df.loc[r,'Top Losers #' + str(i+1) + ' Number of UPs'])
                loser_act.append(new_str(str(df.loc[r,'Top Losers #' + str(i+1) + ' Recent Change Events'])))

                gainers.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Product Type'])
                gainer_ids.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID'])
                gainer_walmarturls.append("https://www.walmart.com/ip/" + str(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID']))
                gainer_urls.append("https://app.ezdia.com/catalog?query=(%20item_id%20EQ%20" + str(df.loc[r,'Top Gainers #' + str(i+1) + ' Item ID']) +"%20)")
                gainer_metrics.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Difference in Order_Quantity'])
                gainer_trends.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Recent Order_Quantity Trend'])
                gainer_downs.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Number of Downs'])
                gainer_ups.append(df.loc[r,'Top Gainers #' + str(i+1) + ' Number of UPs'])
                gainer_act.append(new_str(str(df.loc[r,'Top Gainers #' + str(i+1) + ' Recent Change Events'])))

            losers_data = {'Losers': losers, 'IDs': loser_ids, 'Walmart URLs': loser_walmarturls, 'Crewmachine URLs': loser_urls, 'Revenues': loser_metrics, 'Trends': loser_trends, 'Downs': loser_downs, 'UPs': loser_ups, 'Activity': loser_act}
            df2 = pd.DataFrame(data=losers_data)
            HTML(df2.to_html(render_links=True, escape=False))


            gainers_data = {'Gainers': gainers, 'IDs': gainer_ids, 'Walmart URLs': gainer_walmarturls, 'Crewmachine URLs': gainer_urls, 'Revenues': gainer_metrics, 'Trends': gainer_trends, 'Activity': gainer_act}
            df3 = pd.DataFrame(data=gainers_data)
            HTML(df3.to_html(render_links=True, escape=False))


            dates_and_rev_loss = {'Date': date, 'Loss': loss}
            df4 = pd.DataFrame(data = dates_and_rev_loss)

            dates_and_rev_gain = {'Date': date, 'Gain': gain}
            df5 = pd.DataFrame(data = dates_and_rev_gain)

        data_frames.append(df2)
        data_frames.append(df3)
        data_frames.append(df4)
        data_frames.append(df5)

# %%
get_data(df, value_type, organization_type)
writer = pd.ExcelWriter(file_name + "-simplified.xlsx", engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', startrow= 0, index=False, startcol=0)
wb = writer.book
ws = writer.sheets['Sheet1']
money_fmt = wb.add_format({'num_format': '$#,###.##'})
ws.set_column('B:B', 12, money_fmt)
ws.set_column('H:H', 12, money_fmt)
ws.set_column('A:J', 25)

l = len(dates) + 2

for i in range(0, len(data_frames), 4):

    data_frames[i].to_excel(writer, sheet_name='Sheet1', startrow= l, index=False, startcol=3)
    data_frames[i+1].to_excel(writer, sheet_name='Sheet1', startrow= l + 7, index=False, startcol=3)
    data_frames[i+2].to_excel(writer, sheet_name='Sheet1', startrow= l, index=False, startcol=0)
    data_frames[i+3].to_excel(writer, sheet_name='Sheet1', startrow= l + 7, index=False, startcol=0)
    i += 4
    l += 17
        
writer.save()


