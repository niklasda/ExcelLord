# LineClown (Excel)  
No longer integrates with Line bot and no longer specific to Clowns  

Check the App.config to set input path and output file pattern  
  
Loops all *.xslx in specified folder 
files interpreted as daily monster hunt stats files  
except:
 *BankData.xlsx    which is exported from guild bank, uses the latest if many exists  
 GF *.xlsx         which is GF scores if selected for export  
 GuildStats*.xlsx  which is the manually exported guild stats, calculated kills diff between the 2 latests files  

