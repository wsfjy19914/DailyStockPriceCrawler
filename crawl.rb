# -*- coding: utf-8 -*-

require 'spreadsheet'
require 'rubygems'
require 'nokogiri'
require 'open-uri'

def get_nokogiri_doc(url)
	begin
		html = open(url)
	rescue OpenURI::HTTPError
		return
	end
	Nokogiri::HTML(html.read, nil, 'utf-8')
end

def has_next_page?(doc)
	doc.xpath("//*[@id='main']/ul/a").each {|element|
		return true if element.text == "次へ"
	}
	return false
end

def get_daily_data(doc,code,f)
	doc.xpath("//table[@class='boardFin yjSt marB6']/tr").each {|element|
		# 日付行および株式分割告知を回避
		if element.children[0].text != "日付" && element.children[1][:class] != "through"

			# 日付
			day = element.children[0].text.chomp

			# 始値
			open_price = element.children[1].text.gsub(/,/,'').chomp

			# 高値
			hight_price = element.children[2].text.gsub(/,/,'').chomp

			# 安値
			low_price = element.children[3].text.gsub(/,/,'').chomp

			# 終値
			closing_price = element.children[4].text.gsub(/,/,'').chomp

			# 出来高
			volume = element.children[5].text.gsub(/,/,'').chomp

			#puts "#{code},#{day},#{open_price},#{hight_price},#{low_price},#{closing_price},#{volume}"
                    #データを取る際、バグっている行があるのでそれをこのif文でスキップする
            if low_price != '始値' then
                  #日付を文字列から20140131のようなint型の数値に変換する
                  ypos = day.index('年')#'年'の位置をyposに格納
                  mpos = day.index('月')#'月'の位置をmposに格納
                  dpos = day.index('日')#'日'の位置をdposに格納


                  #年、月、日の数値部分の文字列を切り出して格納
                  yyyy = day[0..(ypos-1)]
                  mm = day[(ypos+1)..(mpos-1)]
                  dd = day[(mpos+1)..(dpos-1)]

                  #1月や、7日などひと桁の数値を表す文字列については、その前に'0'を追加する
                  if mm.length == 1 then
                      mm = "0" + mm
                  end
                  if dd.length == 1 then
                      dd = "0" + dd
                  end


                  #dayにこれまで作った年月日の文字列を結合して格納
                  day = yyyy + mm + dd
                  #結合したdayをint型に変換して格納
                  day = day.to_i

                  f.write("#{code},#{day},#{open_price},#{hight_price},#{low_price},#{closing_price},#{volume}\n")
            end
        end
	}
end



def stockCrawl(start_code)
    begin
        flag = 0
        book = Spreadsheet.open('Shokencode.xls','rb')
        sheet=book.worksheet('Sheet1')
        sheet.each do |row|
            
            # 証券コード
            code=row[1].to_i
            
            if start_code == 0 then
                flag = 1
                start_code = code
                
            elsif flag == 0 && code != start_code then
                #  print code + ","
                #puts flag
                next;
        
            elsif code == start_code then
                flag = 1
            end
            
            #出力ファイル名
            outfilename=(row[1].to_i).to_s+'.csv'
            f = open(outfilename,"w")
            
            start_code = code
            
            puts "Downloading: " + "#{row[0]} #{row[1]}"
        
            # 開始年月日
            start_date='00000000'
            if start_date == '00000000' then
                sy='1900'
                sm='1'
                sd='1'
            else
                sy=start_date[0..3]
                sm=start_date[4..5]
                sd=start_date[6..7]
            end
        
            #終了年月日
            end_date='00000000'
            if end_date == '00000000' then
                day=Time.now
                ey=day.year
                em=day.month
                ed=day.day
            else
                ey=end_date[0..3]
                em=end_date[4..5]
                ed=end_date[6..7]
            end
        
            # 検索日
        
            start_url="http://info.finance.yahoo.co.jp/history/?sy=#{sy}&sm=#{sm}&sd=#{sd}&ey=#{ey}&em=#{em}&ed=#{ed}&tm=d&code=#{code}"
            num=1
            #puts "証券コード,日付,始値,高値,安値,終値,出来高"
            f.write("証券コード,日付,始値,高値,安値,終値,出来高\n")
            loop {
                url = "#{start_url}&p=#{num}"
                doc = get_nokogiri_doc(url)
                get_daily_data(doc,code,f)
                break if !has_next_page?(doc)
                num = num+1
            }
            
        end
    
    rescue
        puts "Error appeared when processing " + start_code.to_s
        puts "Current time: " + Time.new.inspect
        puts "Program will retry after 1 hour automatically! Please wait..."
        sleep(3610)
        stockCrawl(start_code)
    end
end


#puts "Please input the code of the latest stock which has been downloaded:"
#start_code = gets.chomp

if Dir["./*.csv"].length == 0 then
    start_code = 0
else
    latestFile = Dir["./*.csv"][-1]
    start_code = latestFile[(latestFile.index("/") + 1)..(latestFile.index(".csv") - 1)]
end

stockCrawl(start_code.to_i)
