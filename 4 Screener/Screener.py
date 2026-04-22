import pandas as pd
import os
import geopandas as gpd

# import warnings
# warnings.filterwarnings("ignore")

def extract_excel(OutputFileName, SourceFolder_Profiles, gdf_destination, REtechnology, ProfilesFileName, Flag_RunExcelExtraction,SolarMultiple=2.1):
    if Flag_RunExcelExtraction==1:
        
        gdf_destination.rename(columns={"FID":"MSR_ID"} , inplace=True)

        if AnalysisLevel == "country":
            keys = [(ctry, None) for ctry in gdf_destination.CtryName.unique()]
        else:
            keys = list(gdf_destination[["CtryName", "Region"]].drop_duplicates().itertuples(index=False, name=None))


        if REtechnology=="Solar PV":
            pd_SplatReady=pd.DataFrame()
            for ctry, region in keys:

                path = os.path.join(SourceFolder_Profiles, ctry, f"{ctry} {ProfilesFileName}")  
                
                pd_Profiles = pd.read_csv(path)
                LongitudeColumn=pd_Profiles['Longitude']
                LatitudeColumn=pd_Profiles['Latitude']
                pd_Profiles=pd_Profiles.drop(pd_Profiles.iloc[:,1:20].columns,axis=1)
                pd_Profiles['Longitude']=LongitudeColumn
                pd_Profiles['Latitude']=LatitudeColumn

                if region is None:
                    gdf_filtered = gdf_destination[gdf_destination.CtryName == ctry]
                else:
                    gdf_filtered = gdf_destination[(gdf_destination.CtryName == ctry) &(gdf_destination.Region == region)]
                    
                merge = pd.merge(gdf_filtered, pd_Profiles, on="MSR_ID")
                merge = merge.drop(["geometry"], axis=1)
                
                pd_SplatReady= pd_SplatReady.append(merge)
                print(f"Solar appended: {ctry} {region if region else ''}")

            cols = pd_SplatReady.columns.tolist()
            cols=cols
            pd_SplatReady=pd_SplatReady[cols]

            pd_SplatReady.to_csv(OutputFileName)

        if REtechnology=="Wind":
            pd_SplatReady=pd.DataFrame()
            for ctry, region in keys:
                path = os.path.join(SourceFolder_Profiles, ctry, f"{ctry} {ProfilesFileName}")

                pd_Profiles = pd.read_csv(path)
                LongitudeColumn=pd_Profiles['Longitude']
                LatitudeColumn=pd_Profiles['Latitude']
                pd_Profiles=pd_Profiles.drop(pd_Profiles.iloc[:,1:19].columns,axis=1)
                pd_Profiles['Longitude']=LongitudeColumn
                pd_Profiles['Latitude']=LatitudeColumn

                if region is None:
                    gdf_filtered = gdf_destination[gdf_destination.CtryName == ctry]
                else:
                    gdf_filtered = gdf_destination[(gdf_destination.CtryName == ctry) &(gdf_destination.Region == region)]

                merge = pd.merge(gdf_filtered, pd_Profiles, on="MSR_ID")
                merge = merge.drop(["geometry"], axis=1)

                pd_SplatReady= pd_SplatReady.append(merge)
                print(f"Wind appended: {ctry} {region if region else ''}")

            cols = pd_SplatReady.columns.tolist()
            cols=cols
            pd_SplatReady=pd_SplatReady[cols]

            pd_SplatReady.to_csv(OutputFileName)


ControlPathsAndConfigurations=pd.read_excel(r"C:\Users\yulia\Desktop\New repository\Model-Supply-Regions-MSR-Toolset2\4 Screener\ControlFile_Screener.xlsx", sheet_name="PathsAndConfig", index_col=0)

OutputFolder=ControlPathsAndConfigurations.loc["OutputFolder"][0]
if not os.path.isdir(OutputFolder):
    os.makedirs(OutputFolder)


SolarPVSourceFolderCarryingProfiles = ControlPathsAndConfigurations.loc["SolarPVSourceFolderCarryingProfiles"][0]
WindSourceFolderCarryingProfiles = ControlPathsAndConfigurations.loc["WindSourceFolderCarryingProfiles"][0]
SolarPV_ProfilesFileName=ControlPathsAndConfigurations.loc["SolarPV_ProfilesFileName"][0]
WindCF_ProfilesFileName=ControlPathsAndConfigurations.loc["WindCF_ProfilesFileName"][0]
AnalysisLevel = ControlPathsAndConfigurations.loc["AnalysisLevel"][0].strip().lower()

Countries=pd.read_csv(ControlPathsAndConfigurations.loc["FileAddress_CountryNamesList"][0],names=["Ct"])
Regions = pd.read_csv(ControlPathsAndConfigurations.loc["FileAddress_RegionNamesList"][0],encoding="latin1")
RegionBoundariesShapeFile=ControlPathsAndConfigurations.loc["RegionBoundariesShapeFile"][0]
CountryBoundariesShapeFile=ControlPathsAndConfigurations.loc["CountryBoundariesShapeFile"][0]


#Pre-Screen file addresses
SolarPV_PreScreenFile_country=ControlPathsAndConfigurations.loc["SolarPV_PreScreenFile_country"][0]
WindPreScreenFile_country=ControlPathsAndConfigurations.loc["WindPreScreenFile_country"][0]
SolarPV_PreScreenFile_region=ControlPathsAndConfigurations.loc["SolarPV_PreScreenFile_region"][0]
WindPreScreenFile_region=ControlPathsAndConfigurations.loc["WindPreScreenFile_region"][0]
if AnalysisLevel == "country":
    SolarPV_PreScreenFile = SolarPV_PreScreenFile_country
    WindPreScreenFile = WindPreScreenFile_country
elif AnalysisLevel == "region":
    SolarPV_PreScreenFile = SolarPV_PreScreenFile_region
    WindPreScreenFile = WindPreScreenFile_region

# All screening is done after sorting MSRs in descending order of LCOE (supply+transmission+road)
# This script is drafted in such a way that it facilitates the further expansion to any no of screening options

#Current screening options
# 1 Country specific covered area cutoff (Select best MSRs that cover x % of country area)

#read screening options and criteria to apply
ScreeningOptions=pd.read_excel(r"C:\Users\yulia\Desktop\New repository\Model-Supply-Regions-MSR-Toolset2\4 Screener\ControlFile_Screener.xlsx", sheet_name="Select screening option", index_col=0)
CountrySpecificCriteriaSolarPV=pd.read_excel(r"C:\Users\yulia\Desktop\New repository\Model-Supply-Regions-MSR-Toolset2\4 Screener\ControlFile_Screener.xlsx", sheet_name="SolarPV country specific", index_col=0)
CountrySpecificCriteriaWind=pd.read_excel(r"C:\Users\yulia\Desktop\New repository\Model-Supply-Regions-MSR-Toolset2\4 Screener\ControlFile_Screener.xlsx", sheet_name="Wind country specific", index_col=0)



#Read which screening options has been chosen by the user to run (1 means run, 0 means dont run)
ScreeningOptionsSelected= ScreeningOptions[ScreeningOptions['Selection Status']==1].index.astype('int').to_list()
print(ScreeningOptionsSelected)
ExcelExtractionOptionsSelected = ScreeningOptions[ScreeningOptions['Extract in Excel'] == 1].index.astype('int').to_list()
print(ExcelExtractionOptionsSelected)

Flag_RunSolarPV=ControlPathsAndConfigurations.loc["Run code for SolarPV"][0]
Flag_RunWind=ControlPathsAndConfigurations.loc["Run code for Wind"][0]
WindHeight=ControlPathsAndConfigurations.loc["wind hub height in meters"][0]
WindAnnualCFColName = 'CF%sm' % WindHeight
WindAnnualYieldColName = 'Y_GWh%sm' % WindHeight

#Solar MSR Screener
if Flag_RunSolarPV:
    REtechnology="Solar PV"
    gdf= gpd.GeoDataFrame()
    file_to_read=SolarPV_PreScreenFile
    gdf_source = gpd.read_file(file_to_read)
    gdf_source.CF=gdf_source.CF.astype(float)

    # sort in descending order of LCOE
    gdf_source=gdf_source.sort_values(by=['LCOE-MWh'], ascending=True, ignore_index=True)

    for method in ScreeningOptionsSelected:
        print("SolarPV Screening Method %s"%method)
        if method in ExcelExtractionOptionsSelected:
            Flag_RunExcelExtraction=1
            print("Extract Excel")
        else:
            Flag_RunExcelExtraction=0
            print("Not extracting Excel")

        if method==1:
            gdf_destination = gpd.GeoDataFrame()

            if AnalysisLevel == "country":
                iterator = [(ctry.strip(), None) for ctry in Countries["Ct"]]
            else:
                iterator = list(Regions[["Country", "Region"]].drop_duplicates().itertuples(index=False, name=None))
                            

            for Country, Region in iterator:
                Country = Country.strip()          
                CountryKey = Country.replace(" ", "")

                if Region:
                    Region = Region.strip()
                    RegionKey = Region.replace(" ", "")

                if AnalysisLevel == "country":
                    gdf_boundaries = gpd.read_file(CountryBoundariesShapeFile)
                    gdf_selected = gdf_boundaries[gdf_boundaries.name == Country]

                elif AnalysisLevel == "region":
                    gdf_boundaries = gpd.read_file(RegionBoundariesShapeFile)
                    gdf_selected = gdf_boundaries[(gdf_boundaries.name == Region) & (gdf_boundaries.geonunit == Country)] 

                Area_kM2=gdf_selected.to_crs("ESRI:54009").area.iloc[0] / 1000000
                cutoff=(CountrySpecificCriteriaSolarPV.loc[Country][0]/100)*Area_kM2

                if AnalysisLevel == "country":
                    gdf_filtered = gdf_source[gdf_source.CtryName == CountryKey]
                else:
                    gdf_filtered = gdf_source[(gdf_source.CtryName == CountryKey) & (gdf_source.Region == RegionKey)]                

                gdf_filtered['CumAreakM2'] = gdf_filtered.AreakM2.cumsum()

                gdf_filtered = gdf_filtered[gdf_filtered['CumAreakM2'] <= cutoff]


                gdf_destination=gpd.GeoDataFrame(pd.concat([gdf_destination, gdf_filtered]))
                print (f"Processing: {Country}" + (f" - {Region}" if Region else ""))


            gdf_destination=gdf_destination.drop(['CumAreakM2'], axis=1)
            
            if AnalysisLevel == "country":
                suffix = "Country"
            else:
                suffix = "Region"

            shp_path = os.path.join(OutputFolder, f"{REtechnology}_BestMSRs_{suffix}.shp")
            csv_path = os.path.join(OutputFolder, f"{REtechnology}_BestMSRs_{suffix}.csv")

            gdf_destination.to_file(shp_path)
            extract_excel(csv_path, SolarPVSourceFolderCarryingProfiles, gdf_destination, REtechnology, SolarPV_ProfilesFileName,
                          Flag_RunExcelExtraction)

#Wind MSR Screener
if Flag_RunWind:
    REtechnology = "Wind"
    gdf= gpd.GeoDataFrame()
    file_to_read=WindPreScreenFile
    gdf_source = gpd.read_file(file_to_read)
    gdf_source[WindAnnualCFColName] = gdf_source[WindAnnualCFColName].astype(float)

    # sort in descending order of LCOE
    gdf_source=gdf_source.sort_values(by=['LCOE-MWh'], ascending=True, ignore_index=True)

    for method in ScreeningOptionsSelected:
        print("Wind Screening Method %s"%method)
        if method in ExcelExtractionOptionsSelected:
            Flag_RunExcelExtraction=1
            print ("Extract Excel")
        else:
            Flag_RunExcelExtraction=0

        if method==1:
            gdf_destination = gpd.GeoDataFrame()

            if AnalysisLevel == "country":
                iterator = [(ctry.strip(), None) for ctry in Countries["Ct"]]
            else:
                iterator = list(Regions[["Country", "Region"]].drop_duplicates().itertuples(index=False, name=None))
                
            for Country, Region in iterator:
                Country = Country.strip()
                CountryKey = Country.replace(" ", "")
                if Region:
                    Region = Region.strip()
                    RegionKey = Region.replace(" ", "")

                if AnalysisLevel == "country":
                    gdf_boundaries = gpd.read_file(CountryBoundariesShapeFile)
                    gdf_selected = gdf_boundaries[gdf_boundaries.name == Country]

                elif AnalysisLevel == "region":
                    gdf_boundaries = gpd.read_file(RegionBoundariesShapeFile)
                    gdf_selected = gdf_boundaries[(gdf_boundaries.name == Region) & (gdf_boundaries.geonunit == Country)] 

                Area_kM2=gdf_selected.to_crs("ESRI:54009").area.iloc[0] / 1000000
                cutoff=(CountrySpecificCriteriaSolarPV.loc[Country][0]/100)*Area_kM2

                if AnalysisLevel == "country":
                    gdf_filtered = gdf_source[gdf_source.CtryName == CountryKey]
                else:
                    gdf_filtered = gdf_source[(gdf_source.CtryName == CountryKey) & (gdf_source.Region == RegionKey)]                

                gdf_filtered['CumAreakM2'] = gdf_filtered.AreakM2.cumsum()

                gdf_filtered = gdf_filtered[gdf_filtered['CumAreakM2'] <= cutoff]

                gdf_destination=gpd.GeoDataFrame(pd.concat([gdf_destination, gdf_filtered]))
                print (f"Processing: {Country}" + (f" - {Region}" if Region else ""))

            gdf_destination=gdf_destination.drop(['CumAreakM2'], axis=1)
            
            if AnalysisLevel == "country":
                suffix = "Country"
            else:
                suffix = "Region"

            shp_path = os.path.join(OutputFolder, f"{REtechnology}_BestMSRs_{suffix}.shp")
            csv_path = os.path.join(OutputFolder, f"{REtechnology}_BestMSRs_{suffix}.csv")

            gdf_destination.to_file(shp_path)
            extract_excel(
                csv_path, 
                WindSourceFolderCarryingProfiles, 
                gdf_destination, 
                REtechnology, 
                WindCF_ProfilesFileName,
                Flag_RunExcelExtraction
            )
