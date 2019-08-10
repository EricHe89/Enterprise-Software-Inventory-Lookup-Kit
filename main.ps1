# AUTHOR: Henghui He
# This script is created to search particular items in the Excel using PowerShell
# 3 parameters can be inputted as search criteria, and entries can be copied directly 
# from 1 file to another. Please close the Excel before existing the program
# SOURCE_FILE and DESTINATION FILE are both Excels, and need to be exported as .txt
# files first so no actual excels need to be opened during program execution.

################### ADJUST FILE PATH AS NECESSARY ###################
$(SOURCE_FILE)_Path = 'C:\Users\HENGHUI\Desktop\Software Inventory\(SOURCE_FILE).XLSX'
$(SOURCE_FILE) = New-Object -ComObject Excel.Application
$(SOURCE_FILE).Visible = $FALSE
$(SOURCE_FILE)_WorkBook = $(SOURCE_FILE).WorkBooks.Open($(SOURCE_FILE)_Path)
$(SOURCE_FILE)_WorkSheet = $(SOURCE_FILE).WorkSheets.item("Query_Copy")

$(DESTINATION_FILE)_Path = 'C:\Users\HXH\Desktop\Software Inventory\(DESTINATION_FILE).xlsx'
$(DESTINATION_FILE) = New-Object -ComObject Excel.Application
$(DESTINATION_FILE).Visible = $FALSE
$(DESTINATION_FILE)_WorkBook = $(DESTINATION_FILE).Workbooks.Open($(DESTINATION_FILE)_Path)
$(DESTINATION_FILE)_WorkSheet_WithOutEntries = $(DESTINATION_FILE).WorkSheets.Item("(SOURCE_FILE)_entries_not_in_(DESTINATION_FILE)")
$(DESTINATION_FILE)_WorkSheet_WithEntries = $(DESTINATION_FILE).WorkSheets.Item("(DESTINATION_FILE)_(SOURCE_FILE)")	

####################### 1st FUNCTION #######################
function Close_Excel
{	
	$(DESTINATION_FILE)_WorkBook.Save()
	$(DESTINATION_FILE)_WorkBook.Close()
	$(DESTINATION_FILE).Quit()
	$(SOURCE_FILE)_WorkBook.Save()
	$(SOURCE_FILE)_WorkBook.Close()
	$(SOURCE_FILE).Quit()
	
	write-host "`nAll excel sheets are closed now.`n"
}

########################## 2nd FUNCTION #########################################

function Find_It{

        $pattern_1 = $args[0]
        $pattern_2 = $args[1]
        $pattern_3 = $args[2]
        $pattern_4 = $args[3]
 	
        if($pattern_4){
               	 write-host "`nFunction allows only (the first) 3 parameters ..."
               	 write-host "(1)Software Name   (2)Company Name   (3)Version Number`n"
        }elseif($pattern_3)
       {
                echo (select-String -path (DESTINATION_FILE).txt -pattern $pattern_1) > RESULT.txt
               	 write-host "`n"
      
	foreach($LINE in Get-Content -path RESULT.txt) 
	{
	      $substring = ($LINE -split ":")[2]    
	      if($substring -match $pattern_2 -AND $substring -match $pattern_3)
	      {
	               ($LINE -split ":")[0] + "   " +  ($LINE -split ":")[1] + "`t" + ($LINE -split ":")[2]
	       }
                }
	write-host "`n"
        }elseif(!$pattern_3 -AND $pattern_2){
                echo (select-String -path (DESTINATION_FILE).txt -pattern $pattern_1) > RESULT.txt
                 write-host "`n"
	foreach($LINE in Get-Content -path RESULT.txt) 
	{
	      $substring = ($LINE -split ":")[2]    
	      
	      if($substring -match $pattern_2)
	      { 
		($LINE -split ":")[0] + "   " +  ($LINE -split ":")[1] + "`t" + ($LINE -split ":")[2]
	       }
	}
	 write-host "`n"
        }else{
                echo (select-string -path (DESTINATION_FILE).txt -pattern $pattern_1) > RESULT.txt

                foreach($LINE in Get-Content -path RESULT.txt) 
	{
                        ($LINE -split ":")[0] + "   " + ($LINE -split ":")[1] + "`t" + ($LINE -split ":")[2]
                }
        }

}

########################## 3rd FUNCTION #########################################
function SingleCopy_FOUND{
	
$(SOURCE_FILE)_Num = $args[0] #The line number in (SOURCE_FILE) spreadsheet
$(DESTINATION_FILE)_NUM_1 = $args[1] #The 1st line number in (DESTINATION_FILE) found through the "FIND" script
$(DESTINATION_FILE)_NUM_2 = $args[2] #The 2nd line number in (DESTINATION_FILE) found through the "FIND" script
$(DESTINATION_FILE)_NUM_3 = $args[3] #The 3rd line number in (DESTINATION_FILE) found through the "FIND" script
$(DESTINATION_FILE)_NUM_4 = $args[4] #The 4th line number in (DESTINATION_FILE) found through the "FIND" script
$(DESTINATION_FILE)_NUM_5 = $args[5] #The 5th line number in (DESTINATION_FILE) found through the "FIND" script
$(DESTINATION_FILE)_NUM_6 = $args[6] #The 6th line number in (DESTINATION_FILE) found through the "FIND" script
$(DESTINATION_FILE)_NUM_7 = $args[7] #The 7th line number in (DESTINATION_FILE) found through the "FIND" script
$(DESTINATION_FILE)_NUM_8 = $args[8] #The 8th line number in (DESTINATION_FILE) found through the "FIND" script
$(DESTINATION_FILE)_NUM_9 = $args[9] #The 9th line number in (DESTINATION_FILE) found through the "FIND" script
$(DESTINATION_FILE)_NUM_10 = $args[10] #The 10st line number in (DESTINATION_FILE) found through the "FIND" script
$(DESTINATION_FILE)_NUM_11 = $args[11] #The 11th line number in (DESTINATION_FILE) found through the "FIND" script

####### If no (SOURCE_FILE) entries found in (DESTINATION_FILE) ###################################################
if (!$(DESTINATION_FILE)_NUM_1){	

	write-host "`nPlease enter at least 1 (DESTINATION_FILE) Entries, max 10...`n"
}
else{
	if ($(DESTINATION_FILE)_NUM_11){
		write-host "`n...... CAUTION: MORE THAN 10 (DESTINATION_FILE) ENTRIES ENTERED, ONLY FIRST 10 ARE PROCESSED ......"
	}

	######## 1st entry in (DESTINATION_FILE) ###########################################################################################
	$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_1, 5).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 1).Text
	$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_1, 6).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 2).Text
	$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_1, 7).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 3).Text
	$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_1, 8).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 4).Text
	

	######## 2nd entry in (DESTINATION_FILE) ###########################################################################################
	if ($(DESTINATION_FILE)_NUM_2){
		$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_2, 5).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 1).Text
		$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_2, 6).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 2).Text
		$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_2, 7).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 3).Text
		$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_2, 8).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 4).Text
	
		######## 3rd entry in (DESTINATION_FILE) ###########################################################################################
		if ($(DESTINATION_FILE)_NUM_3){
			$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_3, 5).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 1).Text
			$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_3, 6).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 2).Text
			$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_3, 7).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 3).Text
			$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_3, 8).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 4).Text

			######## 4th entry in (DESTINATION_FILE) ###########################################################################################
			if ($(DESTINATION_FILE)_NUM_4){
				$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_4, 5).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 1).Text
				$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_4, 6).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 2).Text
				$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_4, 7).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 3).Text
				$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_4, 8).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 4).Text

				######## 5th entry in (DESTINATION_FILE) ###################################################	########################################	
				if ($(DESTINATION_FILE)_NUM_5){
					$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_5, 5).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 1).Text
					$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_5, 6).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 2).Text
					$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_5, 7).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 3).Text
					$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_5, 8).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 4).Text

					######## 6th entry in (DESTINATION_FILE) ##############################################################################################	
					if ($(DESTINATION_FILE)_NUM_6){
						$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_6, 5).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 1).Text
						$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_6, 6).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 2).Text
						$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_6, 7).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 3).Text
						$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_6, 8).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 4).Text

						######## 7th entry in (DESTINATION_FILE) ###############################################################################################	
						if ($(DESTINATION_FILE)_NUM_7){
							$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_7, 5).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 1).Text
							$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_7, 6).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 2).Text
							$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_7, 7).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 3).Text
							$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_7, 8).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 4).Text
						

							######## 8th entry in (DESTINATION_FILE) ###############################################################################################
							if ($(DESTINATION_FILE)_NUM_8){
								$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_8, 5).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 1).Text
								$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_8, 6).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 2).Text
								$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_8, 7).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 3).Text
								$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_8, 8).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 4).Text		

								######## 9th entry in (DESTINATION_FILE) ###############################################################################################
								if ($(DESTINATION_FILE)_NUM_9){
									$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_9, 5).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 1).Text
									$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_9, 6).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 2).Text
									$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_9, 7).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 3).Text
									$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_9, 8).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 4).Text 

									######## 10st entry in (DESTINATION_FILE) ###############################################################################################
									if ($(DESTINATION_FILE)_NUM_10){
										$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_10, 5).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 1).Text
										$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_10, 6).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 2).Text
										$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_10, 7).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 3).Text
										$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($(DESTINATION_FILE)_NUM_10, 8).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num, 4).Text
									}	
								}
							}	
						}
					}

				}

			}

		}
	}
	write-host "`n...... (SOURCE_FILE) DATA WRITTEN INTO (DESTINATION_FILE)_FOUND......`n"
   }
}

############################# 4th FUNCTION ######################################
function BulkCopy_FOUND
{

	$(SOURCE_FILE)_NUM = $args[0] #The starting line number of the (SOURCE_FILE) entry in spreadsheet
	$(DESTINATION_FILE)_NUM_FIRST = $args[1]
	$(DESTINATION_FILE)_NUM_LAST = $args[2]
	$SPOT_HOLDER = $args[3] # the function only takes 2 parameters 
			# it indicates warning when user inputs >2 paramters	

	######### Warning Messagae######################################
	if ($SPOT_HOLDER -or !$(DESTINATION_FILE)_NUM_LAST)
	{
		write-host "-----------------------------------------------------"
		write-host "`nWARNING: function needs 3 parameters !!!`n"
		write-host "(1) (SOURCE_FILE) entry  (2) First Entry in (DESTINATION_FILE) copy to  (3) Last Entry in (DESTINATION_FILE)`n"
		write-host "Please Verify your data and try again ... `n"
		write-host "-----------------------------------------------------"

	}else{
		For ($index = $(DESTINATION_FILE)_NUM_FIRST; $index -le $(DESTINATION_FILE)_NUM_LAST; $index++){

			$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($index, 5).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_NUM, 1).Text
			$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($index, 6).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_NUM, 2).Text
			$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($index, 7).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_NUM, 3).Text
			$(DESTINATION_FILE)_WorkSheet_WithEntries.Cells.Item($index, 8).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_NUM, 4).Text
		}
	
		write-host "`n...(SOURCE_FILE) Data Copied to (DESTINATION_FILE)_FOUND: $(DESTINATION_FILE)_NUM_FIRST ~ $(DESTINATION_FILE)_NUM_LAST...`n"
	}	
}

##################################### 5th FUNCTION ###################################

function SingleCopy_NOTFOUND{

$(SOURCE_FILE)_Num_0 = $args[0] #The 1st line number in (SOURCE_FILE) spreadsheet
$(SOURCE_FILE)_NUM_1 = $args[1] #The 2nd line number in (SOURCE_FILE) 
$(SOURCE_FILE)_NUM_2 = $args[2] #The 3rd line number in (SOURCE_FILE) 
$(SOURCE_FILE)_NUM_3 = $args[3] #The 4th line number in (SOURCE_FILE) 
$(SOURCE_FILE)_NUM_4 = $args[4] #The 5th line number in (SOURCE_FILE) 
$(SOURCE_FILE)_NUM_5 = $args[5] #The 6th line number in (SOURCE_FILE) 
$(SOURCE_FILE)_NUM_6 = $args[6] #The 7th line number in (SOURCE_FILE)
$(SOURCE_FILE)_NUM_7 = $args[7] #The 8th line number in (SOURCE_FILE) 
$(SOURCE_FILE)_NUM_8 = $args[8] #The 9th line number in (SOURCE_FILE) 
$(SOURCE_FILE)_NUM_9 = $args[9] #The 10th line number in (SOURCE_FILE)
$(SOURCE_FILE)_NUM_10 = $args[10] #The 11th line number in (SOURCE_FILE) which indicates WARNING msg

####### If no (SOURCE_FILE) entries found in (SOURCE_FILE) #############################################################
if(!$(SOURCE_FILE)_NUM_0){
	write-host "`nPlease enter at least 1 (SOURCE_FILE) entery to copy to (DESTINATION_FILE)_NOT_FOUND...`n"

}else{	
	if ($(SOURCE_FILE)_NUM_10){
		write-host "`n...... CAUTION: MORE THAN 10 (SOURCE_FILE) ENTRIES ENTERED, ONLY FIRST 10 ARE PROCESSED ......"}

	######## 1st entry in (SOURCE_FILE) ###########################################################################################
	$LastUsedRow = $(DESTINATION_FILE)_WorkSheet_WithOutEntries.UsedRange.rows.count
	$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 1).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_0, 1).Text
	$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 2).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_0, 2).Text
	$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 3).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_0, 3).Text
	$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 4).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_0, 4).Text


	######## 2nd entry in (SOURCE_FILE) ###########################################################################################
	if ($(SOURCE_FILE)_NUM_1){
	$LastUsedRow = $(DESTINATION_FILE)_WorkSheet_WithOutEntries.UsedRange.rows.count
	$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 1).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_1, 1).Text
	$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 2).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_1, 2).Text
	$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 3).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_1, 3).Text
	$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 4).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_1, 4).Text
	

	######## 3rd entry in (SOURCE_FILE) ###########################################################################################
	if ($(SOURCE_FILE)_NUM_2){
		$LastUsedRow = $(DESTINATION_FILE)_WorkSheet_WithOutEntries.UsedRange.rows.count
		$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 1).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_2, 1).Text
		$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 2).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_2, 2).Text
		$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 3).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_2, 3).Text
		$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 4).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_2, 4).Text
	
		######## 4th entry in (SOURCE_FILE) ###########################################################################################
		if ($(SOURCE_FILE)_NUM_3){
			$LastUsedRow = $(DESTINATION_FILE)_WorkSheet_WithOutEntries.UsedRange.rows.count
			$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 1).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_3, 1).Text
			$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 2).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_3, 2).Text
			$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 3).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_3, 3).Text
			$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 4).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_3, 4).Text

			######## 5th entry in (SOURCE_FILE) ###########################################################################################
			if ($(SOURCE_FILE)_NUM_4){
				$LastUsedRow = $(DESTINATION_FILE)_WorkSheet_WithOutEntries.UsedRange.rows.count
				$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 1).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_4, 1).Text
				$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 2).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_4, 2).Text
				$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 3).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_4, 3).Text
				$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 4).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_4, 4).Text

				######## 6th entry in (SOURCE_FILE) ###################################################	########################################	
				if ($(SOURCE_FILE)_NUM_5){
						$LastUsedRow = $(DESTINATION_FILE)_WorkSheet_WithOutEntries.UsedRange.rows.count
						$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 1).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_5, 1).Text
						$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 2).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_5, 2).Text
						$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 3).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_5, 3).Text
						$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 4).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_5, 4).Text

					######## 7th entry in (SOURCE_FILE) ##############################################################################################	
					if ($(SOURCE_FILE)_NUM_6){
						$LastUsedRow = $(DESTINATION_FILE)_WorkSheet_WithOutEntries.UsedRange.rows.count
						$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 1).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_6, 1).Text
						$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 2).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_6, 2).Text
						$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 3).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_6, 3).Text
						$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 4).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_6, 4).Text

						######## 8th entry in (SOURCE_FILE) ###############################################################################################	
						if ($(SOURCE_FILE)_NUM_7){
							$LastUsedRow = $(DESTINATION_FILE)_WorkSheet_WithOutEntries.UsedRange.rows.count
							$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 1).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_7, 1).Text
							$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 2).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_7, 2).Text
							$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 3).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_7, 3).Text
							$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 4).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_7, 4).Text

							######## 9th entry in (SOURCE_FILE) ###############################################################################################
							if ($(SOURCE_FILE)_NUM_8){
								$LastUsedRow = $(DESTINATION_FILE)_WorkSheet_WithOutEntries.UsedRange.rows.count
								$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 1).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_8, 1).Text
								$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 2).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_8, 2).Text
								$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 3).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_8, 3).Text
								$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 4).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_8, 4).Text		

								######## 10th entry in (SOURCE_FILE) ###############################################################################################
								if ($(SOURCE_FILE)_NUM_9){
									$LastUsedRow = $(DESTINATION_FILE)_WorkSheet_WithOutEntries.UsedRange.rows.count
									$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 1).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_9, 1).Text
									$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 2).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_9, 2).Text
									$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 3).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_9, 3).Text
									$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 4).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($(SOURCE_FILE)_Num_9, 4).Text
						
									}
								}	
							}
						}

					}

				}

			}
		}
	    }
    		write-host "`n...(SOURCE_FILE) Data Written into (DESTINATION_FILE)_NOT_FOUND...`n"
	}
						
}


##################################### 6th FUNCTION ###################################
function BulkCopy_NOTFOUND{

$(SOURCE_FILE)_NUM_FIRST = $args[0] 
$(SOURCE_FILE)_NUM_LAST = $args[1]
$SPOT_HOLDER = $args[2]  

######### Warning Messagae#################################################################
if ($SPOT_HOLDER -or !$args[1]){

	write-host "------------------------------------------------------------"
	write-host "WARNING: function takes 2 parameters !!!`n"
	write-host "(1) first Entry in (SOURCE_FILE) to be copied	 (2) Last Entry in (SOURCE_FILE)`n"
	write-host "Please verify your data and enter again ..."
	write-host "------------------------------------------------------------"

}
else{
	For ($index = $(SOURCE_FILE)_NUM_FIRST; $index -le $(SOURCE_FILE)_NUM_LAST; $index++){
		$LastUsedRow = $(DESTINATION_FILE)_WorkSheet_WithOutEntries.UsedRange.rows.count

		$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 1).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($index, 1).Text
		$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 2).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($index, 2).Text
		$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 3).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($index, 3).Text
		$(DESTINATION_FILE)_WorkSheet_WithOutEntries.Cells.Item($LastUsedRow + 1, 4).Value = $(SOURCE_FILE)_WorkSheet.Cells.Item($index, 4).Text	
	}
	
	write-host "`n... (SOURCE_FILE) Data: $(SOURCE_FILE)_NUM_FIRST ~ $(SOURCE_FILE)_NUM_LAST Copied to (DESTINATION_FILE)_NOT_FOUND ...`n"
    }

}
		                                                       			                                         
write-host "`n**************************************************************"
write-host "        (SILK)  SOFTWARE INVENTORY LOOKUP KIT`n                "
write-host "**************************************************************"
write-host "`  Made by a UC Davis student.  Script-Load Successful.       `n"	
write-host "**************************************************************"
