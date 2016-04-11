require 'spreadsheet'

Spreadsheet.client_encoding = 'UTF-8'


################################
# SET VARIABLES
################################

sworn_a_200 = 0
sworn_a_400 = 0
sworn_a_600 = 0
sworn_a_720 = 0

sworn_s_100 = 0
sworn_s_500 = 0
sworn_s_1500 = 0
sworn_s_2000 = 0
sworn_s_2500 = 0
sworn_s_3000 = 0
sworn_s_3500 = 0

sworn_c_100 = 0
sworn_c_200 = 0
sworn_c_300 = 0
sworn_c_400 = 0
sworn_c_500 = 0

sworn_h_64 = 0
sworn_h_100 = 0
sworn_h_200 = 0
sworn_h_300 = 0
sworn_h_400 = 0
sworn_h_500 = 0
sworn_h_600 = 0
sworn_h_700 = 0

sworn_p_50 = 0
sworn_p_100 = 0
sworn_p_150 = 0
sworn_p_200 = 0
sworn_p_250 = 0

corrections_a_200 = 0
corrections_a_400 = 0
corrections_a_600 = 0
corrections_a_720 = 0

corrections_s_100 = 0
corrections_s_500 = 0
corrections_s_1500 = 0
corrections_s_2000 = 0
corrections_s_2500 = 0
corrections_s_3000 = 0
corrections_s_3500 = 0

corrections_c_100 = 0
corrections_c_200 = 0
corrections_c_300 = 0
corrections_c_400 = 0
corrections_c_500 = 0

corrections_h_64 = 0
corrections_h_100 = 0
corrections_h_200 = 0
corrections_h_300 = 0
corrections_h_400 = 0
corrections_h_500 = 0
corrections_h_600 = 0
corrections_h_700 = 0

corrections_p_50 = 0
corrections_p_100 = 0
corrections_p_150 = 0
corrections_p_200 = 0
corrections_p_250 = 0



################################################################
# LOOP OVER SWORN SPREADSHEETS
################################################################

puts "Starting on sworn spreadsheets..."

Dir.glob("files/sworn/*.xls").each do |file|
  
  puts "\n\nOpening - #{file}\n"
  
  ################################
  # OPEN SPREADSHEET
  ################################
  book = Spreadsheet.open file
  
  ################################
  # OPEN 'TOTALS' WORKSHEET
  ################################
  sheet = book.worksheet 0
    
  ################################
  # ANNUAL
  ################################  
  
  if sheet.row(16)[2]
  
    total_a = sheet.row(16)[2].value.to_i
    
    if total_a > 719 
      sworn_a_720 += 1
    elsif total_a.to_f >= 600
      sworn_a_600 += 1
    elsif total_a.to_f >= 400
      sworn_a_400 += 1
    elsif total_a.to_f >= 200
      sworn_a_200 += 1
    else
    end
    
    puts "  #{total_a} - Annual"
    
  end
  
  ################################
  # SICK
  ################################
    
  if sheet.row(17)[2]
    
    total_s = sheet.row(17)[2].value
  
    if total_s.to_f >= 3500
      sworn_s_3500 += 1
    elsif total_s.to_f >= 3000
      sworn_s_3000 += 1
    elsif total_s.to_f >= 2500
      sworn_s_2500 += 1
    elsif total_s.to_f >= 2000
      sworn_s_2000 += 1
    elsif total_s.to_f >= 1500
      sworn_s_1500 += 1
    elsif total_s.to_f >= 500
      sworn_s_500 += 1
    elsif total_s.to_f >= 100
      sworn_s_100 += 1
    else
    end
    
    puts "  #{total_s} - Sick"
    
  end
  

  
  
  ################################
  # COMPENSATORY
  ################################
    
  if sheet.row(18)[2]
    
    total_c = sheet.row(18)[2].value
  
    if total_c.to_f >= 500
      sworn_c_500 += 1
    elsif total_c.to_f >= 400
      sworn_c_400 += 1
    elsif total_c.to_f >= 300
      sworn_c_300 += 1
    elsif total_c.to_f >= 200
      sworn_c_200 += 1
    elsif total_c.to_f >= 100
      sworn_c_100 += 1
    else
    end
    
    puts "  #{total_c} - Comp"
    
  end
  

  
  
  ################################
  # HOLIDAY
  ################################
  
  if sheet.row(19)[2]
    
    total_h = sheet.row(19)[2].value
    
    if total_h.to_f >= 700
      sworn_h_700 += 1
    elsif total_h.to_f >= 600
      sworn_h_600 += 1
    elsif total_h.to_f >= 500
      sworn_h_500 += 1
    elsif total_h.to_f >= 400
      sworn_h_400 += 1
    elsif total_h.to_f >= 300
      sworn_h_300 += 1
    elsif total_h.to_f >= 200
      sworn_h_200 += 1
    elsif total_h.to_f >= 100
      sworn_h_100 += 1
    elsif total_h.to_f >= 64
      sworn_h_64 += 1
    else
    end
    
    puts "  #{total_h} - Holiday"
    
  end
  
  
  

  
  ################################
  # PERSONAL
  ################################
      
  if sheet.row(13)[2]
            
    total_p = sheet.row(13)[2]
  
    if total_p.to_f >= 250
      sworn_p_250 += 1
    elsif total_p.to_f >= 200
      sworn_p_200 += 1
    elsif total_p.to_f >= 150
      sworn_p_150 += 1
    elsif total_p.to_f >= 100
      sworn_p_100 += 1
    elsif total_p.to_f >= 50
      sworn_p_50 += 1
    else
    end
    
    puts "  #{total_p.to_f} - Personal"
    
  end

end


################################################################
# LOOP OVER CORRECTIONS SPREADSHEETS
################################################################

Dir.glob("files/corrections/*.xls").each do |file|
  
  puts "\n\nOpening - #{file}\n"
  
  ################################
  # OPEN SPREADSHEET
  ################################
  book = Spreadsheet.open file
  
  ################################
  # OPEN 'TOTALS' WORKSHEET
  ################################
  sheet = book.worksheet 0
    
  ################################
  # ANNUAL
  ################################  
  
  if sheet.row(16)[2]
  
    total_a = sheet.row(16)[2].value.to_i
    
    if total_a >= 719 
      corrections_a_720 += 1
    elsif total_a.to_f >= 600
      corrections_a_600 += 1
    elsif total_a.to_f >= 400
      corrections_a_400 += 1
    elsif total_a.to_f >= 200
      corrections_a_200 += 1
    else
    end
    
    puts "  #{total_a} - Annual"
    
  end
  
  ################################
  # SICK
  ################################
    
  if sheet.row(17)[2]
    
    total_s = sheet.row(17)[2].value
  
    if total_s.to_f >= 3500
      corrections_s_3500 += 1
    elsif total_s.to_f >= 3000
      corrections_s_3000 += 1
    elsif total_s.to_f >= 2500
      corrections_s_2500 += 1
    elsif total_s.to_f >= 2000
      corrections_s_2000 += 1
    elsif total_s.to_f >= 1500
      corrections_s_1500 += 1
    elsif total_s.to_f >= 500
      corrections_s_500 += 1
    elsif total_s.to_f >= 100
      corrections_s_100 += 1
    else
    end
    
    puts "  #{total_s} - Sick"
    
  end
  

  
  
  ################################
  # COMPENSATORY
  ################################
    
  if sheet.row(18)[2]
    
    total_c = sheet.row(18)[2].value
  
    if total_c.to_f >= 500
      corrections_c_500 += 1
    elsif total_c.to_f >= 400
      corrections_c_400 += 1
    elsif total_c.to_f >= 300
      corrections_c_300 += 1
    elsif total_c.to_f >= 200
      corrections_c_200 += 1
    elsif total_c.to_f >= 100
      corrections_c_100 += 1
    else
    end
    
    puts "  #{total_c} - Comp"
    
  end
  

  
  
  ################################
  # HOLIDAY
  ################################
  
  if sheet.row(19)[2]
    
    total_h = sheet.row(19)[2].value ? sheet.row(19)[2].value : 0
    
    if total_h.to_f >= 700
      corrections_h_700 += 1
    elsif total_h.to_f >= 600
      corrections_h_600 += 1
    elsif total_h.to_f >= 500
      corrections_h_500 += 1
    elsif total_h.to_f >= 400
      corrections_h_400 += 1
    elsif total_h.to_f >= 300
      corrections_h_300 += 1
    elsif total_h.to_f >= 200
      corrections_h_200 += 1
    elsif total_h.to_f >= 100
      corrections_h_100 += 1
    elsif total_h.to_f >= 64
      corrections_h_64 += 1
    else
    end
    
    puts "  #{total_h} - Holiday"
    
  end
  
  
  

  
  ################################
  # PERSONAL
  ################################
      
  if sheet.row(13)[2]
            
    total_p = sheet.row(13)[2]
  
    if total_p.to_f >= 250
      corrections_p_250 += 1
    elsif total_p.to_f >= 200
      corrections_p_200 += 1
    elsif total_p.to_f >= 150
      corrections_p_150 += 1
    elsif total_p.to_f >= 100
      corrections_p_100 += 1
    elsif total_p.to_f >= 50
      corrections_p_50 += 1
    else
    end
    
    puts "  #{total_p.to_f} - Personal"
    
  end
    
    
end


################################
# SAVE STATS TO SPREADSHEET
################################

stats_book = Spreadsheet.open './leave_stats_template.xls'

sworn_stats_sheet = stats_book.worksheet 0
corrections_stats_sheet = stats_book.worksheet 1

puts "Saving stats to spreadsheet\n\n"

sworn_stats_sheet.row(1)[3] = sworn_a_200
sworn_stats_sheet.row(2)[3] = sworn_a_400
sworn_stats_sheet.row(3)[3] = sworn_a_600
sworn_stats_sheet.row(4)[3] = sworn_a_720

sworn_stats_sheet.row(1)[1] = sworn_s_100
sworn_stats_sheet.row(2)[1] = sworn_s_500
sworn_stats_sheet.row(3)[1] = sworn_s_1500
sworn_stats_sheet.row(4)[1] = sworn_s_2000
sworn_stats_sheet.row(5)[1] = sworn_s_2500
sworn_stats_sheet.row(6)[1] = sworn_s_3000
sworn_stats_sheet.row(7)[1] = sworn_s_3500

sworn_stats_sheet.row(1)[5] = sworn_c_100
sworn_stats_sheet.row(2)[5] = sworn_c_200
sworn_stats_sheet.row(3)[5] = sworn_c_300
sworn_stats_sheet.row(4)[5] = sworn_c_400
sworn_stats_sheet.row(5)[5] = sworn_c_500

sworn_stats_sheet.row(1)[7] = sworn_h_64
sworn_stats_sheet.row(2)[7] = sworn_h_100
sworn_stats_sheet.row(3)[7] = sworn_h_200
sworn_stats_sheet.row(4)[7] = sworn_h_300
sworn_stats_sheet.row(5)[7] = sworn_h_400
sworn_stats_sheet.row(6)[7] = sworn_h_500
sworn_stats_sheet.row(7)[7] = sworn_h_600
sworn_stats_sheet.row(8)[7] = sworn_h_700

sworn_stats_sheet.row(1)[9] = sworn_p_50
sworn_stats_sheet.row(2)[9] = sworn_p_100
sworn_stats_sheet.row(3)[9] = sworn_p_150
sworn_stats_sheet.row(4)[9] = sworn_p_200
sworn_stats_sheet.row(5)[9] = sworn_p_250

corrections_stats_sheet.row(1)[3] = corrections_a_200
corrections_stats_sheet.row(2)[3] = corrections_a_400
corrections_stats_sheet.row(3)[3] = corrections_a_600
corrections_stats_sheet.row(4)[3] = corrections_a_720

corrections_stats_sheet.row(1)[1] = corrections_s_100
corrections_stats_sheet.row(2)[1] = corrections_s_500
corrections_stats_sheet.row(3)[1] = corrections_s_1500
corrections_stats_sheet.row(4)[1] = corrections_s_2000
corrections_stats_sheet.row(5)[1] = corrections_s_2500
corrections_stats_sheet.row(6)[1] = corrections_s_3000
corrections_stats_sheet.row(7)[1] = corrections_s_3500

corrections_stats_sheet.row(1)[5] = corrections_c_100
corrections_stats_sheet.row(2)[5] = corrections_c_200
corrections_stats_sheet.row(3)[5] = corrections_c_300
corrections_stats_sheet.row(4)[5] = corrections_c_400
corrections_stats_sheet.row(5)[5] = corrections_c_500

corrections_stats_sheet.row(1)[7] = corrections_h_64
corrections_stats_sheet.row(2)[7] = corrections_h_100
corrections_stats_sheet.row(3)[7] = corrections_h_200
corrections_stats_sheet.row(4)[7] = corrections_h_300
corrections_stats_sheet.row(5)[7] = corrections_h_400
corrections_stats_sheet.row(6)[7] = corrections_h_500
corrections_stats_sheet.row(7)[7] = corrections_h_600
corrections_stats_sheet.row(8)[7] = corrections_h_700

corrections_stats_sheet.row(1)[9] = corrections_p_50
corrections_stats_sheet.row(2)[9] = corrections_p_100
corrections_stats_sheet.row(3)[9] = corrections_p_150
corrections_stats_sheet.row(4)[9] = corrections_p_200
corrections_stats_sheet.row(5)[9] = corrections_p_250

stats_book.write './leave_stats.xls'

puts "DONE!!!"