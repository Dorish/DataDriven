# To change this template, choose Tools | Templates
# and open the template in the editor.

$:.unshift File.dirname(__FILE__)
require 'unity_componets.rb'

module Unity_actions
  include Unity_componets
  
  # - Navigate to a specific page
  def navigate_to(page_link)
    send(page_link).click
  end

  def clickbtn(button_name)
    while (send(button_name).enabled? == false)
      sleep 0.5
    end
    send(button_name).click
  end
  
  def click_no_wait_btn(button_name)
    while (send(button_name).enabled? == false)
      sleep 0.5
    end
    send(button_name).click_no_wait
  end

  def select_combobox(name,text)
    send(name).select(text)
  end

  def set_textbox(name,text)
    send(name).set(text)
  end

  def set_checkbox(chkbox, state)
    send(chkbox).send state
  end

  def waitsave(sec)
    sleep sec.to_i # wait the values being updated
  end

  def verify_result(name, param ,sheet, row)
    if send(name).type == 'checkbox'
      actu_value = checkbox(send(name))
    elsif send(name).id =~ /enum/
      actu_value = send(name).selected_options.to_s
    elsif send(name).id =~ /str/
      actu_value = send(name).text
    else
      puts "not define yet."
    end
    sheet.Range("L#{row}")['Value'] = actu_value
    
    if actu_value == param
      sheet.Range("M#{row}")['Value'] = 'Pass'
    else
      sheet.Range("M#{row}")['Value'] = 'Fail'
    end
  end
end

