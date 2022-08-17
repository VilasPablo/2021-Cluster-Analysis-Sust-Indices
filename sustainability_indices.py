import pandas as pd
import numpy as np
import datetime
import ast

from sqlalchemy import create_engine
from scipy import stats
from dateutil import relativedelta

import warnings
warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)

# In order to assing the "*" of significance
def significance(p_val):
    if p_val <=0.05 and p_val > 0.01 :
        significance = "*"
    elif p_val <=0.01 and p_val > 0.001:
        significance = "**"
    elif p_val <=0.001 :
        significance = "**"
    else:
        significance=""
    return significance


class EquityIndex(object):
    
    column_date = "MonthlyDate"
    column_company_id = "IsinCode"
    variables_analyzed = ['EsgScore', 'CombinedScore', 'MarketValueUsd'] # 'ESG Controversies Score'
    
    def __init__(self, sql_engine, index_name, index_market, index_type, start_date, end_date):     
        
        self.name = index_name
        self.geographic_area = [index_market]
        self.type = index_type   
        # select the index composition data
        sql_query = """
            SELECT "IsinCode", "MonthlyDate", "CompanyWeight"
            FROM "EquityIndicesConstituents"
            WHERE ("IndexTicker" = '%s') AND ("MonthlyDate" BETWEEN '%s' AND '%s')
            ORDER BY "MonthlyDate"
            """ % (self.name, datetime.date(*start_date), datetime.date(*end_date) )
        
        self.df_index = pd.read_sql(sql_query, sql_engine) 
        self.df_index[self.name] = "constituent" # Create a column used when merge index information with ESG score
        
    
    # We create the column ("last_month") ->"in" identifies the companies that will dissapear next month
    def get_index_exclusions(self):   
        def companies_next_month(company_id): # if the company_id is not in the index next month then exclusion 
            if company_id in companies_ids_compare:
                return np.nan 
            else:
                return "in" 
        
        analyzed_dates = sorted( list(set(self.df_index[self.column_date])) ) # unique dates in the index sorted
        for fecha in analyzed_dates[:-1]: # the last date can not be anlyze
            fecha_compare = analyzed_dates[analyzed_dates.index(fecha)+1] # next month date
            companies_ids_compare = list( set(self.df_index[self.df_index[self.column_date] == fecha_compare][
                self.column_company_id]) )# ids of companies that are in the index next month
            
            # if the ids are not next month then is the last month of the company in the index -->assing company "in" label
            self.df_index.loc[self.df_index[self.column_date] == fecha, self.name + "LastMonth"] = self.df_index[
                self.df_index[self.column_date] == fecha][self.column_company_id].apply(companies_next_month) 
            
    # We create a column ("first_month") in identifies the companies that is its first month in the index       
    def get_index_inclusions(self):   
        def companies_last_month(company_id): # if the id was not last month then is inclusion 
            if company_id in companies_ids_compare:
                return np.nan
            else:
                return "in" 

        analyzed_dates = sorted( list(set(self.df_index[self.column_date])) ) # unique dates in the index sorted
        for fecha in analyzed_dates[1:]: # the first date can not be anlyze
            fecha_compare = analyzed_dates[analyzed_dates.index(fecha)-1] # last month date
            companies_ids_compare = list( set(self.df_index[self.df_index[self.column_date] == fecha_compare][
                self.column_company_id]) )# ids of companies that were in the index last month
            # if the ids were not last month then is the first month of the company in the index -->assing company "in" label
            self.df_index.loc[self.df_index[self.column_date] == fecha, self.name + "FirstMonth"] = self.df_index[
                self.df_index[self.column_date] == fecha][self.column_company_id].apply(companies_last_month) 
            
            
    def get_rankpct_exclusions(self, EquityIndexUniverse, restriction, old_methodology= False):
        
        self.df_rankpct_exclusions = pd.DataFrame()
        # exclusion values can be measured at two different times
        if restriction == "out":
            on_column = "FirstMonth"            
        elif restriction == "in":
            on_column = "LastMonth"
        else:
            raise  ValueError("restriction should be: in or out")
            
        self.exclusions_column =on_column
        self.exclusions_restriction = restriction 
            
        # to obtain the rank percentile we focus on exclusions + companies in the index (constituents)
        self.df_rankpct_exclusions = EquityIndexUniverse.df_universe.loc[ (EquityIndexUniverse.df_universe[
            self.name + on_column]==restriction)|(EquityIndexUniverse.df_universe[self.name] == "constituent"),
                EquityIndex.variables_analyzed+ [self.name, self.name + "FirstMonth", self.name + "LastMonth"] +[
                    EquityIndexUniverse.column_company_country ] ] # [self.name "FirstMonth" ... select only these columns]
        
        # We employ the same criteria that in the first article version
        if old_methodology == True:        
            self.df_rankpct_exclusions = self.df_rankpct_exclusions[(self.df_rankpct_exclusions[
                self.name + "FirstMonth"]!="in") ]
            
        # We only analyze the maintenances in the same dates that there exist exclusions
        dates = list(self.df_rankpct_exclusions.loc[self.df_rankpct_exclusions[self.name + on_column] == 
                                               restriction].index.get_level_values(1) )
        self.df_rankpct_exclusions = self.df_rankpct_exclusions.loc[
            self.df_rankpct_exclusions.index.get_level_values(1).isin(dates)]

        
        # get the percentile rank 
        for variable in EquityIndex.variables_analyzed:
            self.df_rankpct_exclusions["Rankpct" + variable] = self.df_rankpct_exclusions.groupby(
                [EquityIndexUniverse.primary_key[1]])[ variable].rank( method ="min")
            self.df_rankpct_exclusions["Rankpct" + variable] = self.df_rankpct_exclusions.groupby([
                EquityIndexUniverse.primary_key[1]])["Rankpct" + variable].apply(lambda x: (x-1)/(x.count()-1) )
           
            
    def get_rankpct_inclusions(self, EquityIndexUniverse, restriction, old_methodology= False):
        
        self.df_rankpct_inclusions = pd.DataFrame()
        # inclusion values can be measured at two different times
        if restriction == "in" :
            on_column = "FirstMonth"         
        elif restriction == "out":
            on_column = "LastMonth"
        else:
            raise  ValueError("restriction should be: in or out")
            
        self.inclusions_column =on_column
        self.inclusions_restriction = restriction 
            
        # to obtain the rank percentile we focus on each inclusions + companies in the same geogrphic are and not in the index
        if self.geographic_area[0] == "Global":
            self.df_rankpct_inclusions = EquityIndexUniverse.df_universe.loc[ 
                  (EquityIndexUniverse.df_universe[self.name] !="constituent")| (
                     EquityIndexUniverse.df_universe[self.name + on_column] == restriction), EquityIndex.variables_analyzed +[
                      self.name, self.name + "FirstMonth", self.name + "LastMonth"] +[
                     EquityIndexUniverse.column_company_country ] ]     
        else:
            self.df_rankpct_inclusions = EquityIndexUniverse.df_universe.loc[ 
                 (EquityIndexUniverse.df_universe[EquityIndexUniverse.column_company_country].isin(self.geographic_area)) & (
                     EquityIndexUniverse.df_universe[self.name] !="constituent")| (
                     EquityIndexUniverse.df_universe[self.name + on_column] == restriction), EquityIndex.variables_analyzed +[
                      self.name, self.name + "FirstMonth", self.name + "LastMonth"] +[
                     EquityIndexUniverse.column_company_country ] ]
        
     # We employ the same criteria that in the first article version
        if old_methodology == True:  
            self.df_rankpct_inclusions = self.df_rankpct_inclusions[ (self.df_rankpct_inclusions[
                self.name+ 'FirstMonth']!='out') ]
            
        # We only analyze the maintenances in the same dates that there exist exclusions
        dates = list(self.df_rankpct_inclusions.loc[self.df_rankpct_inclusions[self.name + on_column] == 
                                               restriction].index.get_level_values(1) )
        self.df_rankpct_inclusions = self.df_rankpct_inclusions.loc[
            self.df_rankpct_inclusions.index.get_level_values(1).isin(dates)]


        for variable in EquityIndex.variables_analyzed:
            self.df_rankpct_inclusions["Rankpct" + variable] = self.df_rankpct_inclusions.groupby(
                [EquityIndexUniverse.primary_key[1]])[variable].rank( method ="min")
            self.df_rankpct_inclusions["Rankpct" + variable] = self.df_rankpct_inclusions.groupby([
                EquityIndexUniverse.primary_key[1]])["Rankpct" + variable].apply(lambda x: (x-1)/(x.count()-1) )

                        
class EquityIndexUniverse (object):
    
    primary_key = ("IsinCode", "MonthlyDate")
    
    
    def __init__(self, sql_engine, country_classification):
        import pandas as pd
        
        self.column_company_country = country_classification
        self.df_universe = pd.read_sql('CompaniesMonthlyData', sql_engine)
        self.df_universe.set_index(list(self.primary_key), inplace=True)
        
        
    # merge the composition information of the index with the monthly ESG scores (universe ESG scores)
    def get_info_from_index(self, EquityIndex):
        df_index = EquityIndex.df_index.drop_duplicates(subset=[EquityIndex.column_company_id, EquityIndex.column_date])
        # we merge the information by the index so the name of both indexes level have to be the same
        df_index = df_index.set_index([EquityIndex.column_company_id, EquityIndex.column_date])
        df_index.index.set_names([*self.primary_key], inplace = True)
        # we are only interested in three columns
        df_index = df_index.loc[:, [EquityIndex.name, EquityIndex.name + "LastMonth", EquityIndex.name +"FirstMonth"]]
        
    # if the name of the index is in the dataframe then combine first else merge// Only merge the isins that are in universe
        if EquityIndex.name in self.df_universe.columns:
            df_index = df_index.loc[df_index.index.intersection([*(self.df_universe.index)]) ]
            self.df_universe = self.df_universe.combine_first(df_index) 
        else:
            self.df_universe = self.df_universe.merge(df_index, how="left", on= list(self.primary_key) ) 
        
    # the first month that company was out of the index
    def get_last_month_out(self, index_name):
        # We create a dataframe with the date and isin that indicates the first month that company was out of the index            
        # We select the date ans ISIN that indentifiy the first time that a company was in the index
        last_month_out =self.df_universe.reset_index()
        last_month_out =last_month_out.loc[ last_month_out[index_name + "FirstMonth"]== "in", 
                                           list(self.primary_key) ]
        # We subtract one month to those dates
        last_month_out[self.primary_key[1] ]=last_month_out[self.primary_key[1] ].apply(
            lambda x: x +relativedelta.relativedelta(months=-1))
        # We create a column (index.name + "_last_month") indicating the last month that a company was out of the index
        last_month_out[index_name + "LastMonth"]= "out"
        last_month_out.set_index( list(self.primary_key), inplace= True)
        # only the dates and companis with score in df_universe (our esg scores database)
        last_month_out = last_month_out.loc[last_month_out.index.intersection([*(self.df_universe.index)])]
        # We update the column  (index.name + "_last_month") of df_universe
        self.df_universe= self.df_universe.combine_first(last_month_out)
        
    def get_first_month_out(self, index_name):
        # We create a dataframe with the date and isin that indicates the first month that a company was out of the index
        
        # We select the date ans ISIN that indentifiy the last time that a company was in the index
        first_month_out =self.df_universe.reset_index()
        first_month_out =first_month_out.loc[ first_month_out[index_name + "LastMonth"] == "in", 
                                             list(self.primary_key) ]
        # We add one month to those dates
        first_month_out[self.primary_key[1]]=first_month_out[self.primary_key[1]].apply(
            lambda x: x +relativedelta.relativedelta(months=+1))
        # We create a column (index.name + "_first_month") indicating the first month that a company was out of the index
        first_month_out[index_name+  "FirstMonth"]= "out"
        first_month_out.set_index(list(self.primary_key), inplace= True)        
        # only the dates with score in df_universe (our esg scores database)
        first_month_out = first_month_out.loc[first_month_out.index.intersection([*(self.df_universe.index)])]
        # We update the column  (index.name + "_first_month") of df_universe
        self.df_universe= self.df_universe.combine_first(first_month_out)

        
class TestAgainstGroup (object):
    
    def __init__(self, sql_engine, nature):
                
        # to see if the test is inclusions against universe or exclusions egainst maintenances
        self.nature = nature
        if nature == "Inclusions":
            self.other = "Universe"
            self.percentile = 0.8
            self.percentile_label = "(2)H0: μ1 ≤" + str(self.percentile)  # the ≤ changes between inclusions and exclusions 
        elif nature == "Exclusions":
            self.other = "Maintenances"
            self.percentile = 0.2
            self.percentile_label ="(2)H0: μ1 ≥" + str(self.percentile)
        else:
            raise ValueError("nature should be: Inclusions or Exclusions")
        # Read the table from database    
        self.df_test = pd.read_sql(self.nature + 'AgainstGroup', con = sql_engine)
        self.df_test.set_index(['level_0', 'level_1', 'level_2'], inplace = True) 

    # restart the dataframe df_test_inclusions 
    def reset(self, sql_engine):
        # We create the index with the headers that will in the table
        column_index = pd.MultiIndex.from_product([EquityIndex.variables_analyzed,[self.nature, self.other],["mean"]]).append(
            pd.MultiIndex.from_product([EquityIndex.variables_analyzed,['Test'],["(1)H0: μ1 = μ2",self.percentile_label]])) 
        column_index= column_index.append(pd.MultiIndex.from_product([["obs"], [self.nature, self.other],["#"]]) )
        column_index= column_index.append(pd.MultiIndex.from_product([[""], [""],["type"]]) )
        self.df_test = pd.DataFrame(index= column_index).reindex(EquityIndex.variables_analyzed+ ["obs", ""],
                                                                       level=0, axis=0)        
        self.df_test.to_sql(self.nature + 'AgainstGroup', con = sql_engine, if_exists = "replace", method = 'multi')
        self.df_test= pd.read_sql(self.nature + 'AgainstGroup', con = sql_engine)
        self.df_test.set_index(['level_0', 'level_1', 'level_2'], inplace = True)
          
        # We apply the test of the paper to see differences between inclusions and universe
    def get_test(self, EquityIndex, start_date, end_date): 

       # We select the correct df_rankpct     
        if self.nature == "Inclusions":
            df_rankpct = EquityIndex.df_rankpct_inclusions.copy()
            on_column = EquityIndex.inclusions_column
            restriction = EquityIndex.inclusions_restriction
        if self.nature == "Exclusions":
            df_rankpct = EquityIndex.df_rankpct_exclusions.copy()
            on_column = EquityIndex.exclusions_column
            restriction = EquityIndex.exclusions_restriction       

        # we select the dates that interest us
        df_rankpct = df_rankpct[ df_rankpct.index.get_level_values(1)>= datetime.datetime(*start_date)].copy()
        df_rankpct = df_rankpct[df_rankpct.index.get_level_values(1)<= datetime.datetime(*end_date)].copy()

        for variable in EquityIndex.variables_analyzed:
    # filer to get the values of each variable we are interested in (inclusions & universe // exclusions & maintenances)
            nature_vector = df_rankpct.loc[df_rankpct[ EquityIndex.name + on_column]== restriction, ["Rankpct" + variable]]
            nature_vector = nature_vector["Rankpct" + variable].tolist() # exclusions / inclusions vector
            other_vector = df_rankpct.loc[df_rankpct[EquityIndex.name + on_column] != restriction, ["Rankpct"+variable]]
            other_vector = other_vector["Rankpct" + variable].tolist()

            # nature mean --> inclusion or exclusion mean
            self.df_test.at[(variable, self.nature,'mean'), EquityIndex.name]=str(
                np.round(np.mean(nature_vector),3))
            # other mean --> universe or maintenances
            self.df_test.at[(variable, self.other,'mean'), EquityIndex.name]=str(
                np.round(np.mean(other_vector),3))
            # T-test of mean differences between two samples: inclusions group = universe group
            variance_test = stats.bartlett(nature_vector, other_vector)
            equal_var = True if variance_test[1] > 0.05 else False
            mean_two_stat, mean_two_p = stats.ttest_ind(nature_vector, other_vector, equal_var=equal_var)
            self.df_test.at[(variable, 'Test', "(1)H0: μ1 = μ2"), EquityIndex.name] =(
                                                    str(np.round(mean_two_stat,1)) + significance(mean_two_p) )
            # T-test of one sample: universe_mean > or < than self.percentile
            mean_one_stat, mean_one_p = stats.ttest_1samp(nature_vector, self.percentile) # (x-mu)/(S/raiz(N-1))
            mean_one_p = stats.t.cdf(mean_one_stat, df= len(nature_vector) - 1)
            mean_one_p = 1- mean_one_p if self.nature == "Inclusions" else mean_one_p
            self.df_test.at[(variable, 'Test', self.percentile_label), EquityIndex.name] =(
                                                    str(np.round(mean_one_stat,1)) + significance(mean_one_p) )
        # number of inclusions 
        self.df_test.at[('obs', self.nature, "#"), EquityIndex.name] = str(len(nature_vector))
        self.df_test.at[('obs', self.other, "#"), EquityIndex.name] = str(len(other_vector)) 
        # type of index 
        self.df_test.at[("", "","type"), EquityIndex.name] = EquityIndex.type
        
    # save the df into the database
    def save(self, sql_engine): 
        self.df_test.to_sql(self.nature + 'AgainstGroup', con = sql_engine, if_exists = "replace", method = 'multi',
                           chunksize=100000)
        

class TestAgainstVariable (object):
    
    def __init__(self, sql_engine, nature, variable_against):
        
        self.variable_against = variable_against
        self.nature = nature
        if nature == "Inclusions":
            pass
        elif nature == "Exclusions":
            pass
        else:
            raise ValueError("nature should be: Inclusions or Exclusions")
        
        self.df_test = pd.read_sql(self.nature + 'AgainstVariable', con = sql_engine)
        self.df_test.set_index(['level_0', 'level_1', 'level_2'], inplace = True)

    def reset(self, sql_engine):

        # We create the index with the headers that will in the table
        column_index = pd.MultiIndex.from_product([EquityIndex.variables_analyzed,[self.nature],["var", "mean"]]).append(
          pd.MultiIndex.from_product([ list(set(EquityIndex.variables_analyzed) - set([self.variable_against])), 
                                      ['Test'],["(1)H0:σ2S = σ2CSP", "(2)H0:μS = μCSP"]])) 
        column_index= column_index.append(pd.MultiIndex.from_product([["obs"], [self.nature],["#"]]) )
        column_index= column_index.append(pd.MultiIndex.from_product([[""], [""],["type"]]) )
        self.df_test = pd.DataFrame(index= column_index).reindex( [self.variable_against] +list(
            set(EquityIndex.variables_analyzed) - set([self.variable_against])) + ["obs", ""] , level=0, axis=0)      
        
        self.df_test.to_sql(self.nature + 'AgainstVariable', con = sql_engine, if_exists = "replace", 
                                         method = 'multi')
        self.df_test= pd.read_sql(self.nature + 'AgainstVariable', con = sql_engine)
        self.df_test.set_index(['level_0', 'level_1', 'level_2'], inplace = True)
          
    # We apply the test of the paper to see differences in inclsuions between size and csp
    def get_test(self, EquityIndex, start_date, end_date):    
     
       # We select the correct df_rankpct     
        if self.nature == "Inclusions":
            df_rankpct = EquityIndex.df_rankpct_inclusions.copy()
            on_column = EquityIndex.inclusions_column
            restriction = EquityIndex.inclusions_restriction
        if self.nature == "Exclusions":
            df_rankpct = EquityIndex.df_rankpct_exclusions.copy()
            on_column = EquityIndex.exclusions_column
            restriction = EquityIndex.exclusions_restriction 

        # we select the dates that interest us
        df_rankpct = df_rankpct[ ( df_rankpct.index.get_level_values(1)>= datetime.datetime(*start_date) ) &(
            df_rankpct.index.get_level_values(1)<= datetime.datetime(*end_date) )].copy()

        # the variable against
        against_vector=df_rankpct.loc[df_rankpct[EquityIndex.name + on_column]== restriction,[
            "Rankpct" + self.variable_against]]
        against_vector = against_vector["Rankpct" + self.variable_against].tolist()
        # variable_against mean and var
        self.df_test.at[(self.variable_against, self.nature, "mean"), EquityIndex.name] = (
        str(np.round(np.mean(against_vector),3)) )
        self.df_test.at[(self.variable_against, self.nature, "var"), EquityIndex.name] = (
        str(np.round(np.var(against_vector), 3)) )
        # observations --> number of inclusions
        self.df_test.at[('obs', self.nature, "#"), EquityIndex.name] = str(len(against_vector))
        
        # size position against csp position
        for variable in EquityIndex.variables_analyzed:
            if self.variable_against == variable:
                continue # we are not interested in compare for example size against size 
            # filer to get the values of each variable we are interested in
            other_vector=df_rankpct.loc[df_rankpct[EquityIndex.name +on_column]== restriction,["Rankpct" + variable]]
            other_vector = other_vector["Rankpct" + variable].tolist()
            # other_vector mean and var
            self.df_test.at[(variable,self.nature,'mean'),EquityIndex.name]=str(np.round(np.mean(other_vector),3))
            self.df_test.at[(variable,self.nature,'var'),EquityIndex.name]=str(np.round(np.var(other_vector),3))
            # Test of  different variances between two samples:
            variance_stat, variance_p = stats.bartlett(against_vector, other_vector)
            self.df_test.at[(variable, 'Test', "(1)H0:σ2S = σ2CSP"), EquityIndex.name] =(
                                                    str(np.round(variance_stat,1)) + significance(variance_p) )
            # T-test of mean differences between two samples: e.g. csp_inclusions_position = size_inclusions_position
            equal_var = True if variance_p > 0.05 else False
            mean_stat, mean_p = stats.ttest_ind(against_vector, other_vector, equal_var=equal_var)
            self.df_test.at[(variable, 'Test', "(2)H0:μS = μCSP"), EquityIndex.name] =(
                                                    str(np.round(mean_stat,1)) + significance(mean_p) )
        # we add information about the type of index
        self.df_test.at[('', '', 'type'), EquityIndex.name] = EquityIndex.type

    def save(self, sql_engine): 
        self.df_test.to_sql(self.nature + 'AgainstVariable', con = sql_engine, if_exists = "replace", method = 'multi',
                           chunksize = 100000)  

        
class YearlyPositionVar(object):
    
    def __init__(self, sql_engine):
        self.df_position = pd.read_sql('YearlyPositionVar', sql_engine)
        self.df_position.set_index(['IndexTicker', 'Year'], inplace=True)
        
    def get_yearly_position(self, EquityIndex, start_date, end_date):
                    
        start_date = datetime.datetime(*start_date)
        end_date = datetime.datetime(*end_date)
        # get the exclusion postion between dates 
        df_rankpct_exclusions = EquityIndex.df_rankpct_exclusions.copy()
        df_rankpct_exclusions = df_rankpct_exclusions.loc[(df_rankpct_exclusions.index.get_level_values(1) >= start_date) &
                                                (df_rankpct_exclusions.index.get_level_values(1) <= end_date) &
                                                (df_rankpct_exclusions[EquityIndex.name +EquityIndex.exclusions_column]==
                                                 EquityIndex.exclusions_restriction)]
        # get the inclusion postion between dates 
        df_rankpct_inclusions = EquityIndex.df_rankpct_inclusions.copy()
        df_rankpct_inclusions = df_rankpct_inclusions.loc[(df_rankpct_inclusions.index.get_level_values(1) >= start_date) &
                                                (df_rankpct_inclusions.index.get_level_values(1) <= end_date) &
                                                (df_rankpct_inclusions[EquityIndex.name +EquityIndex.inclusions_column]==
                                                 EquityIndex.inclusions_restriction)]
                       
        for year in range(start_date.year, end_date.year + 1):
            # filter to get the year exclusions
            df_rankpct_year_exclusions = df_rankpct_exclusions.loc[
                                            (df_rankpct_exclusions.index.get_level_values(1).year == year)]   
            df_rankpct_year_inclusions = df_rankpct_inclusions.loc[
                                            (df_rankpct_inclusions.index.get_level_values(1).year == year)]
            # get the yearly mean and variance of inclusions and exclusions of each index
            self.get_statistics(df_rankpct_year_exclusions, 'Exclusions', year, EquityIndex)
            self.get_statistics(df_rankpct_year_exclusions, 'Inclusions', year, EquityIndex)
            
        # get the position and var for all period
        self.get_statistics(df_rankpct_exclusions, 'Exclusions', 'All period', EquityIndex)
        self.get_statistics(df_rankpct_inclusions, 'Inclusions', 'All period', EquityIndex)
          
    def get_statistics(self, df, nature, year, EquityIndex):
            # get the mean and variance of the rank percnetile variables
            for variable in ['Rankpct'+ variable for variable in EquityIndex.variables_analyzed]:
                self.df_position.at[(EquityIndex.name, str(year)), variable[7:]+ nature +'Mean'] = df[variable].mean()
                self.df_position.at[(EquityIndex.name, str(year)), variable[7:]+ nature +'Var'] = df[variable].var()
                
            self.df_position.at[(EquityIndex.name, str(year)), "Num" + nature] = df[variable].count()             
            self.df_position.at[(EquityIndex.name, str(year)),'IndexMarket']= EquityIndex.geographic_area[0]
            self.df_position.at[(EquityIndex.name, str(year)),'IndexType'] = EquityIndex.type
          
    def reset(self, sql_engine):

        # We create the index with the headers that will in the table
        self.df_test = pd.DataFrame(columns =['IndexTicker', 'Year'] )      
        self.df_test.to_sql('YearlyPositionVar', con = sql_engine, if_exists = "replace", method = 'multi', index=False)
        self.df_test= pd.read_sql('YearlyPositionVar', con = sql_engine)
        self.df_test.set_index(['IndexTicker', 'Year'], inplace = True)    
    
    def save(self, sql_engine):
        self.df_position.to_sql('YearlyPositionVar', sql_engine, if_exists = 'replace', index = True)
