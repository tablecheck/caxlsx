require 'tc_helper'

class TestTheme < Test::Unit::TestCase
  def setup
    @theme = Axlsx::Theme.new
  end

  def test_pn
    assert_equal(Axlsx::THEME_PN, @theme.pn)
  end

  def test_to_xml_string_returns_valid_xml
    xml = @theme.to_xml_string

    # Basic structure checks
    assert xml.include?('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    assert xml.include?('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">')
    assert xml.include?('</a:theme>')

    # Required sections
    assert xml.include?('<a:themeElements>')
    assert xml.include?('<a:clrScheme name="Office">')
    assert xml.include?('<a:fontScheme name="Office">')
    assert xml.include?('<a:fmtScheme name="Office">')
    assert xml.include?('<a:objectDefaults>')
    assert xml.include?('<a:extraClrSchemeLst/>')
  end

  def test_to_xml_string_with_string_parameter
    str = ''
    result = @theme.to_xml_string(str)

    # Should return the same string object that was passed in
    assert_same(str, result)
    assert !str.empty?
    assert str.include?('<a:theme')
  end

  def test_color_scheme_elements
    xml = @theme.to_xml_string

    # Check for required color scheme elements
    assert xml.include?('<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>')
    assert xml.include?('<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>')
    assert xml.include?('<a:dk2><a:srgbClr val="1F497D"/></a:dk2>')
    assert xml.include?('<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>')

    # Check accent colors
    (1..6).each do |i|
      assert xml.include?("<a:accent#{i}>")
    end

    # Check hyperlink colors
    assert xml.include?('<a:hlink><a:srgbClr val="0000FF"/></a:hlink>')
    assert xml.include?('<a:folHlink><a:srgbClr val="800080"/></a:folHlink>')
  end

  def test_font_scheme_elements
    xml = @theme.to_xml_string

    # Check for major and minor fonts
    assert xml.include?('<a:majorFont>')
    assert xml.include?('<a:latin typeface="Cambria"/>')
    assert xml.include?('<a:minorFont>')
    assert xml.include?('<a:latin typeface="Calibri"/>')
  end

  def test_format_scheme_elements
    xml = @theme.to_xml_string

    # Check for format scheme sections
    assert xml.include?('<a:fillStyleLst>')
    assert xml.include?('<a:lnStyleLst>')
    assert xml.include?('<a:effectStyleLst>')
    assert xml.include?('<a:bgFillStyleLst>')
  end

  def test_object_defaults
    xml = @theme.to_xml_string

    # Check for object defaults
    assert xml.include?('<a:spDef>')
    assert xml.include?('<a:lnDef>')
    assert xml.include?('<a:spPr/>')
    assert xml.include?('<a:bodyPr/>')
    assert xml.include?('<a:lstStyle/>')
  end

  def test_3d_elements_present
    xml = @theme.to_xml_string

    # Check for 3D elements that are crucial for Excel compatibility
    assert xml.include?('<a:scene3d>')
    assert xml.include?('<a:camera prst="orthographicFront">')
    assert xml.include?('<a:lightRig rig="threePt" dir="t">')
    assert xml.include?('<a:sp3d>')
    assert xml.include?('<a:bevelT w="63500" h="25400"/>')
  end

  def test_xml_is_single_line_with_no_whitespace_padding
    xml = @theme.to_xml_string

    # XML should not contain extra whitespace or newlines
    assert !xml.include?("\n")
    assert !xml.include?("  ") # No double spaces
  end
end
