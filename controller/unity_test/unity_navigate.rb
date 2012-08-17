# To change this template, choose Tools | Templates
# and open the template in the editor.

class Unity_Navigate

 
  #  - Tab area frameset abstration
  def tab
    frame_text = self.redirect {$ie.show_frames}
    if frame_text =~ /tabArea/ then $ie.frame(:name, 'tabArea')
    else $ie
    end
  end

 
  def unity_config(unity_tab)
    tab.element(:id, unity_tab)
  end

  # - Navigation link frameset abstration
  def nav
    frame_text = self.redirect {$ie.show_frames}
    if frame_text =~ /mainFrameSet/ then $ie.frame(:id, 'main').frame(:id, 'navigationFrame')
    else $ie.frame(:id, 'navigationFrame')
    end
  end

  # - Navigate to a special page
  def navigate_node(navigate_node)
    nav.link(:text, navigate_node)
  end

  #  - returns true or false if the web page under test has a frame named
  #  - <code>frame_name</code>
  def has_frame?(frame_name)
    frame_text = self.redirect {$ie.show_frames}
    !frame_text.match("#{frame_name}").nil?
  end

  #   - buttons
  #   - check boxes
  #   - combo boxes
  #   - text fields
  def det
    if has_frame?('main') then  $ie.frame(:id, 'main').frame(:id, 'rframeset').frame(:id, 'detailArea')
    else $ie.frame(:id, 'detailArea')
    end
  end

  # button

  def click(button_name)
      #$ie.frame(:id, 'detailArea').button(:id, 'editButton')
      det.button(:id, button_name)
  end

  # file filed
  def set_filefield(form_name, field_name)
    $ie.form(:name, form_name).file_field(:name, field_name)
  end

  # text field
  def set_text_value(form_name, field_id)
      det.form(:name, form_name).text_field(:name, field_id)
  end
  
  # combo box
  def select_combo(form_name, select_name)
      det.form(:name, form_name).select_list(:name, select_name)
  end

  # check box
  def set_check_value(form_name,checkbox_id)
    det.form(:name, form_name).checkbox(:name, checkbox_id)
  end

    #This method is used to redirect stdout to a string
  def redirect
    orig_defout = $defout
    $stdout = StringIO.new
    yield
    $stdout.string
  ensure
    $stdout = orig_defout
  end

  def login(site,user,pswd)
    conn_to = 'Connect to '+ site
    Thread.new{
      thread_cnt = Thread.list.size
      sleep 1 #This sleep is critical, timing may need to be adjusted
      Watir.autoit.WinWait(conn_to)
      Watir.autoit.WinActivate(conn_to)
      Watir.autoit.Send(user)
      Watir.autoit.Send('{TAB}')
      Watir.autoit.Send(pswd)
      popup('Windows Internet Explorer','OK') #launch thread for alert popup
      Watir.autoit.Send('{ENTER}')
    }
  end

end

