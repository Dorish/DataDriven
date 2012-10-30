# To change this template, choose Tools | Templates
# and open the template in the editor.

module Unity_componets

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

  #This method is used to redirect stdout to a string
  def redirect
    orig_defout = $defout
    $stdout = StringIO.new
    yield
    $stdout.string
  ensure
    $stdout = orig_defout
  end
  
  #  - returns true or false if the web page under test has a frame named
  def has_frame?(frame_name)
    frame_text = self.redirect {$ie.show_frames}
    !frame_text.match("#{frame_name}").nil?
  end

  def det
    if has_frame?('main') then  $ie.frame(:id, 'main').frame(:id, 'rframeset').frame(:id, 'detailArea')
    else $ie.frame(:id, 'detailArea')
    end
  end


  # - Navigation link frameset abstration
  def nav
    frame_text = self.redirect {$ie.show_frames}
    if frame_text =~ /mainFrameSet/ then $ie.frame(:id, 'main').frame(:id, 'navigationFrame')
    else $ie.frame(:id, 'navigationFrame')
    end
  end

  def configuration
    nav.link(:id, 'report164160')
  end
  
  def system
    nav.link(:id, 'report164170')
  end
  
  def time_service_settings
    nav.link(:id,'report163930')
  end

  def edit_btn
    det.button(:id, 'editButton')
  end
  def save_btn
    det.button(:id, 'submitButton')
  end

  def cancel_btn
    det.button(:id,'cancelButton')
  end

  def ex_time_source
    det.select_list(:id,'enum103')
  end

  def ntp_server
    det.text_field(:id, 'str33')
  end

  def ntp_sync_rate
    det.select_list(:id,'enum34')
  end

  def time_zone
    det.select_list(:id,'enum35')
  end

  def auto_sync_to_device
    det.checkbox(:id, 'chkbx31')
  end

  def auto_sync_rate
    det.select_list(:id,'enum123')
  end
end
