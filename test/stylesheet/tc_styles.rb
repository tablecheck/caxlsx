require 'tc_helper.rb'

class TestStyles < Test::Unit::TestCase
  def setup
    @styles = Axlsx::Styles.new
  end

  def teardown
  end

  def test_valid_document
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@styles.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.size == 0)
  end

  def test_add_style_border_hash
    border_count = @styles.borders.size
    @styles.add_style :border => { :style => :thin, :color => "FFFF0000" }
    assert_equal(@styles.borders.size, border_count + 1)
    assert_equal(@styles.borders.last.prs.last.color.rgb, "FFFF0000")
    assert_raise(ArgumentError) { @styles.add_style :border => { :color => "FFFF0000" } }
    assert_equal @styles.borders.last.prs.size, 4
  end

  def test_add_style_border_array
    prev_border_count = @styles.borders.size

    borders_array = [
      { :style => :thin, :color => "DDDDDD" },
      { :edges => [:top], :style => :thin, :color => "000000" },
      { :edges => [:bottom], :style => :thick, :color => "FF0000" },
      { :edges => [:left], :style => :dotted, :color => "FFFF00" },
      { :edges => [:right], :style => :dashed, :color => "FFFFFF" },
      { :style => :thick, :color => "CCCCCC" }
    ]

    @styles.add_style(border: borders_array)

    assert_equal(@styles.borders.size, (prev_border_count + 1))

    current_border = @styles.borders.last

    borders_array.each do |b_opts|
      if b_opts[:edges]
        border_pr = current_border.prs.detect { |x| x.name == b_opts[:edges].first }
        assert_equal(border_pr.color.rgb, "FF#{b_opts[:color]}")
      end
    end
  end

  def test_add_style_border_edges
    @styles.add_style :border => { :style => :thin, :color => "0000FFFF", :edges => [:top, :bottom] }
    parts = @styles.borders.last.prs
    parts.each { |pr| assert_equal(pr.color.rgb, "0000FFFF", "Style is applied to #{pr.name} properly") }
    assert((parts.map { |pr| pr.name.to_s }.sort && ['bottom', 'top']).size == 2, "specify two edges, and you get two border prs")
  end

  def test_do_not_alter_options_in_add_style
    # This should test all options, but for now - just the bits that we know caused some pain
    options = { :border => { :style => :thin, :color => "FF000000" } }
    @styles.add_style options
    assert_equal options[:border][:style], :thin, 'thin style is stil in option'
    assert_equal options[:border][:color], "FF000000", 'color is stil in option'
  end

  def test_parse_num_fmt
    f_code = { :format_code => "YYYY/MM" }
    num_fmt = { :num_fmt => 5 }
    assert_equal(@styles.parse_num_fmt_options, nil, 'noop if neither :format_code or :num_fmt exist')
    max = @styles.numFmts.map { |nf| nf.numFmtId }.max
    @styles.parse_num_fmt_options(f_code)
    assert_equal(@styles.numFmts.last.numFmtId, max + 1, "new numfmts gets next available id")
    assert(@styles.parse_num_fmt_options(num_fmt).is_a?(Integer), "Should return the provided num_fmt if not dxf")
    assert(@styles.parse_num_fmt_options(num_fmt.merge({ :type => :dxf })).is_a?(Axlsx::NumFmt), "Makes a new NumFmt if dxf")
  end

  def test_parse_border_options_hash_required_keys
    assert_raise(ArgumentError, "Require color key") { @styles.parse_border_options(:border => { :style => :thin }) }
    assert_raise(ArgumentError, "Require style key") { @styles.parse_border_options(:border => { :color => "FF0d0d0d" }) }
    assert_nothing_raised { @styles.parse_border_options(:border => { :style => :thin, :color => "FF000000" }) }
  end

  def test_parse_border_basic_options
    b_opts = { :border => { :diagonalUp => 1, :edges => [:left, :right], :color => "FFDADADA", :style => :thick } }
    b = @styles.parse_border_options b_opts
    assert(b.is_a? Integer)
    assert_equal(@styles.parse_border_options(b_opts.merge({ :type => :dxf })).class, Axlsx::Border)
    assert(@styles.borders.last.diagonalUp == 1, "border options are passed in to the initializer")
  end

  def test_parse_border_options_edges
    b_opts = { :border => { :diagonalUp => 1, :edges => [:left, :right], :color => "FFDADADA", :style => :thick } }
    @styles.parse_border_options b_opts
    b = @styles.borders.last
    left = b.prs.select { |bpr| bpr.name == :left }[0]
    right = b.prs.select { |bpr| bpr.name == :right }[0]
    top = b.prs.select { |bpr| bpr.name == :top }[0]
    bottom = b.prs.select { |bpr| bpr.name == :bottom }[0]
    assert_equal(top, nil, "unspecified top edge should not be created")
    assert_equal(bottom, nil, "unspecified bottom edge should not be created")
    assert(left.is_a?(Axlsx::BorderPr), "specified left edge is set")
    assert(right.is_a?(Axlsx::BorderPr), "specified right edge is set")
    assert_equal(left.style, right.style, "edge parts have the same style")
    assert_equal(left.style, :thick, "the style is THICK")
    assert_equal(right.color.rgb, left.color.rgb, "edge parts are colors are the same")
    assert_equal(right.color.rgb, "FFDADADA", "edge color rgb is correct")
  end

  def test_parse_border_options_noop
    assert_equal(@styles.parse_border_options({}), nil, "noop if the border key is not in options")
  end

  def test_parse_border_options_integer_xf
    assert_equal(@styles.parse_border_options(:border => 1), 1)
    assert_raise(ArgumentError, "unknown border index") { @styles.parse_border_options(:border => 100) }
  end

  def test_parse_border_options_integer_dxf
    b_opts = { :border => { :edges => [:left, :right], :color => "FFFFFFFF", :style => :thick } }
    b = @styles.parse_border_options(b_opts)
    b2 = @styles.parse_border_options(:border => b, :type => :dxf)
    assert(b2.is_a?(Axlsx::Border), "Cloned existing border object")
  end

  def test_parse_alignment_options
    assert_equal(@styles.parse_alignment_options {}, nil, "noop if :alignment is not set")
    assert(@styles.parse_alignment_options(:alignment => {}).is_a?(Axlsx::CellAlignment))
  end

  def test_parse_font_using_defaults
    original = @styles.fonts.first
    @styles.add_style :b => 1, :sz => 99
    created = @styles.fonts.last
    original_attributes = Axlsx.instance_values_for(original)
    assert_equal(1, created.b)
    assert_equal(99, created.sz)
    copied = original_attributes.reject { |key, value| %w(b sz).include? key }
    instance_vals = Axlsx.instance_values_for(created)
    copied.each do |key, value|
      assert_equal(instance_vals[key], value)
    end
  end

  def test_parse_font_options
    options = {
      :fg_color => "FF050505",
      :sz => 20,
      :b => 1,
      :i => 1,
      :u => :single,
      :strike => 1,
      :outline => 1,
      :shadow => 1,
      :charset => 9,
      :family => 1,
      :font_name => "woot font"
    }
    assert_equal(@styles.parse_font_options {}, nil, "noop if no font keys are set")
    assert(@styles.parse_font_options(:b => 1).is_a?(Integer), "return index of font if not :dxf type")
    assert_equal(@styles.parse_font_options(:b => 1, :type => :dxf).class, Axlsx::Font, "return font object if :dxf type")

    f = @styles.parse_font_options(options.merge(:type => :dxf))
    color = options.delete(:fg_color)
    options[:name] = options.delete(:font_name)
    options.each do |key, value|
      assert_equal(f.send(key), value, "assert that #{key} was parsed")
    end
    assert_equal(f.color.rgb, color)
  end

  def test_parse_fill_options
    assert_equal(@styles.parse_fill_options {}, nil, "noop if no fill keys are set")
    assert(@styles.parse_fill_options(:bg_color => "DE").is_a?(Integer), "return index of fill if not :dxf type")
    assert_equal(@styles.parse_fill_options(:bg_color => "DE", :type => :dxf).class, Axlsx::Fill, "return fill object if :dxf type")
    f = @styles.parse_fill_options(:bg_color => "DE", :type => :dxf)
    assert(f.fill_type.bgColor.rgb == "FFDEDEDE")
  end

  def test_parse_protection_options
    assert_equal(@styles.parse_protection_options {}, nil, "noop if no protection keys are set")
    assert_equal(@styles.parse_protection_options(:hidden => 1).class, Axlsx::CellProtection, "creates a new cell protection object")
  end

  def test_add_style
    fill_count = @styles.fills.size
    font_count = @styles.fonts.size
    xf_count = @styles.cellXfs.size

    @styles.add_style :bg_color => "FF000000", :fg_color => "FFFFFFFF", :sz => 13, :num_fmt => Axlsx::NUM_FMT_PERCENT, :alignment => { :horizontal => :left }, :border => Axlsx::STYLE_THIN_BORDER, :hidden => true, :locked => true
    assert_equal(@styles.fills.size, fill_count + 1)
    assert_equal(@styles.fonts.size, font_count + 1)
    assert_equal(@styles.cellXfs.size, xf_count + 1)
    xf = @styles.cellXfs.last
    assert_equal(xf.fillId, (@styles.fills.size - 1), "points to the last created fill")
    assert_equal(@styles.fills.last.fill_type.fgColor.rgb, "FF000000", "fill created with color")

    assert_equal(xf.fontId, (@styles.fonts.size - 1), "points to the last created font")
    assert_equal(@styles.fonts.last.sz, 13, "font sz applied")
    assert_equal(@styles.fonts.last.color.rgb, "FFFFFFFF", "font color applied")

    assert_equal(xf.borderId, Axlsx::STYLE_THIN_BORDER, "border id is set")
    assert_equal(xf.numFmtId, Axlsx::NUM_FMT_PERCENT, "number format id is set")

    assert(xf.alignment.is_a?(Axlsx::CellAlignment), "alignment was created")
    assert_equal(xf.alignment.horizontal, :left, "horizontal alignment applied")
    assert_equal(xf.protection.hidden, true, "hidden protection set")
    assert_equal(xf.protection.locked, true, "cell locking set")
    assert_raise(ArgumentError, "should reject invalid borderId") { @styles.add_style :border => 2 }

    assert_equal(xf.applyProtection, true, "protection applied")
    assert_equal(xf.applyBorder, true, "border applied")
    assert_equal(xf.applyNumberFormat, true, "number format applied")
    assert_equal(xf.applyAlignment, true, "alignment applied")
  end

  def test_basic_add_style_dxf
    border_count = @styles.borders.size
    @styles.add_style :border => { :style => :thin, :color => "FFFF0000" }, :type => :dxf
    assert_equal(@styles.borders.size, border_count, "styles borders not affected")
    assert_equal(@styles.dxfs.last.border.prs.last.color.rgb, "FFFF0000")
    assert_raise(ArgumentError) { @styles.add_style :border => { :color => "FFFF0000" }, :type => :dxf }
    assert_equal @styles.borders.last.prs.size, 4
  end

  def test_add_style_dxf
    fill_count = @styles.fills.size
    font_count = @styles.fonts.size
    dxf_count = @styles.dxfs.size

    style = @styles.add_style :bg_color => "FF000000", :fg_color => "FFFFFFFF", :sz => 13, :alignment => { :horizontal => :left }, :border => { :style => :thin, :color => "FFFF0000" }, :hidden => true, :locked => true, :type => :dxf
    assert_equal(@styles.dxfs.size, dxf_count + 1)
    assert_equal(0, style, "returns the zero-based dxfId")

    dxf = @styles.dxfs.last
    assert_equal(@styles.dxfs.last.fill.fill_type.bgColor.rgb, "FF000000", "fill created with color")

    assert_equal(font_count, (@styles.fonts.size), "font not created under styles")
    assert_equal(fill_count, (@styles.fills.size), "fill not created under styles")

    assert(dxf.border.is_a?(Axlsx::Border), "border is set")
    assert_equal(nil, dxf.numFmt, "number format is not set")

    assert(dxf.alignment.is_a?(Axlsx::CellAlignment), "alignment was created")
    assert_equal(dxf.alignment.horizontal, :left, "horizontal alignment applied")
    assert_equal(dxf.protection.hidden, true, "hidden protection set")
    assert_equal(dxf.protection.locked, true, "cell locking set")
    assert_raise(ArgumentError, "should reject invalid borderId") { @styles.add_style :border => 3 }
  end

  def test_multiple_dxf
    # add a second style
    style = @styles.add_style :bg_color => "00000000", :fg_color => "FFFFFFFF", :sz => 13, :alignment => { :horizontal => :left }, :border => { :style => :thin, :color => "FFFF0000" }, :hidden => true, :locked => true, :type => :dxf
    assert_equal(0, style, "returns the first dxfId")
    style = @styles.add_style :bg_color => "FF000000", :fg_color => "FFFFFFFF", :sz => 13, :alignment => { :horizontal => :left }, :border => { :style => :thin, :color => "FFFF0000" }, :hidden => true, :locked => true, :type => :dxf
    assert_equal(1, style, "returns the second dxfId")
  end

  def test_valid_document_with_font_options
    font_options = {
      :fg_color => "FF050505",
      :sz => 20,
      :b => 1,
      :i => 1,
      :u => :single,
      :strike => 1,
      :outline => 1,
      :shadow => 1,
      :charset => 9,
      :family => 1,
      :font_name => "woot font"
    }
    @styles.add_style font_options

    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@styles.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.size == 0)
  end

  def test_border_top_without_border_regression
    ### https://github.com/axlsx-styler-gem/axlsx_styler/issues/31

    borders = {
      top: { style: :double, color: '0000FF' },
      right: { style: :thick, color: 'FF0000' },
      bottom: { style: :double, color: '0000FF' },
      left: { style: :thick, color: 'FF0000' }
    }

    borders.each do |edge, b_opts|
      @styles.add_style("border_#{edge}".to_sym => b_opts)

      current_border = @styles.borders.last

      border_pr = current_border.prs.detect { |x| x.name == edge }
      assert_equal(border_pr.color.rgb, "FF#{b_opts[:color]}")
    end
  end
end
