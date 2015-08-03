class FruitsController < ApplicationController
  before_action :set_fruit, only: [:show, :edit, :update, :destroy]

  # GET /fruits
  # GET /fruits.json
  def index
    @fruits = Fruit.all.includes(:season)
  end

  # GET /fruits/1
  # GET /fruits/1.json
  def show
  end

  # GET /fruits/new
  def new
    @fruit = Fruit.new
    @seasons = Season.all
  end

  # GET /fruits/1/edit
  def edit
    @seasons = Season.all
  end

  # POST /fruits
  # POST /fruits.json
  def create
    @fruit = Fruit.new(fruit_params)

    respond_to do |format|
      if @fruit.save
        format.html { redirect_to @fruit, notice: 'Fruit was successfully created.' }
        format.json { render :show, status: :created, location: @fruit }
      else
        format.html { render :new }
        format.json { render json: @fruit.errors, status: :unprocessable_entity }
      end
    end
  end

  # PATCH/PUT /fruits/1
  # PATCH/PUT /fruits/1.json
  def update
    respond_to do |format|
      if @fruit.update(fruit_params)
        format.html { redirect_to @fruit, notice: 'Fruit was successfully updated.' }
        format.json { render :show, status: :ok, location: @fruit }
      else
        format.html { render :edit }
        format.json { render json: @fruit.errors, status: :unprocessable_entity }
      end
    end
  end

  # DELETE /fruits/1
  # DELETE /fruits/1.json
  def destroy
    @fruit.destroy
    respond_to do |format|
      format.html { redirect_to fruits_url, notice: 'Fruit was successfully destroyed.' }
      format.json { head :no_content }
    end
  end

  def download_ss
    # 季節毎にシート分けしたエクセル book をつくろう(Spreadsheet編)
    book = Spreadsheet::Workbook.new
    Season.order(:id).each do |s|
      # シート作成
      sheet = book.create_worksheet name: s.name
      # 一行目はヘッダー
      sheet.update_row(0, "ID", "名前")
      # データ行
      idx = 1
      Fruit.where(season_id: s.id).order('updated_at DESC').each do |fruit|
        sheet.row(idx).push(fruit.id)
        sheet.row(idx).push(fruit.name)
        idx += 1
      end
    end

    filename = "fruits_by_season#{Time.now.strftime('%Y%m%d%H%M%S')}_ss.xls"

    require 'stringio'
    data = StringIO.new ''
    book.write data
    send_data(
      data.string.bytes.to_a.pack("C*"),
      filename: filename,
      type: "application/excel")
  end

  def upload_ss
    book = Spreadsheet.open params[:file].tempfile.path

    Fruit.transaction do
      book.worksheets.each do |ws|
        season = Season.find_by_name!(ws.name)
        ws.rows.each_with_index do |row, idx|
          next if idx < 1
          Fruit.create(season: season, name: row[0])
        end
      end
    end

    redirect_to fruits_path
  end

  def download_rx
    # 季節毎にシート分けしたエクセル book をつくろう(rubyXL編)
    book = RubyXL::Workbook.new
    Season.order(:id).each_with_index do |s, idx|
      # シート作成
      if idx <1
        sheet = book[0]
        sheet.sheet_name = s.name
      else
        sheet = book.add_worksheet(s.name)
      end
      # 一行目はヘッダー
      sheet.add_cell(0, 0, 'ID')   #A1
      sheet.add_cell(0, 1, '名前') #B1
      # データ行
      idx = 1
      Fruit.where(season_id: s.id).order('updated_at DESC').each do |fruit|
        sheet.add_cell(idx, 0, fruit.id)
        sheet.add_cell(idx, 1, fruit.name)
        idx += 1
      end
    end

    filename = "fruits_by_season#{Time.now.strftime('%Y%m%d%H%M%S')}_rx.xlsx"

    send_data(
      book.stream.read,
      filename: filename,
      type: "application/excel")
  end

  def upload_rx
    book = RubyXL::Parser.parse params[:file].tempfile.path

    Fruit.transaction do
      (0...book.worksheets.size).each do |s_idx|
        # シート毎の処理
        sheet = book.worksheets[s_idx]

        season = Season.find_by_name!(sheet.sheet_name)

        (0...sheet.sheet_data.size).each do |d_idx|
          # シート内の行毎の処理
          cell = sheet.sheet_data.rows[d_idx][0]
          next if cell.value.blank?
          Fruit.create(season: season, name: cell.value)
        end
      end
    end

    redirect_to fruits_path
  end

  private
    # Use callbacks to share common setup or constraints between actions.
    def set_fruit
      @fruit = Fruit.find(params[:id])
    end

    # Never trust parameters from the scary internet, only allow the white list through.
    def fruit_params
      params.require(:fruit).permit(:name, :season_id)
    end
end
