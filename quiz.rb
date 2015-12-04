# -*- coding:utf-8 -*-

require 'fox16'
require 'stringio'
require 'spreadsheet'

include Fox

# puts '启动问卷应用 ... '

QUESTIONS = <<"EOF"
当有人对我不善时，我常试图谅解他。
当我一旦决定了某事时，我很不乐意改变想法。
被他人接受对我很重要。
我认为不能原谅愚蠢的错误。
当某事让我感到无聊时，我常用手指敲击东西。
请求得到我需要的东西对我很难。
在他人面前我不愿意表现自己的弱点。
我感觉我被别人利用。
一旦我着手某项工作，就会把它做彻底。
耐心不是我的强项。
EOF


qs =StringIO.new(QUESTIONS)

$book = Spreadsheet::Workbook.new
$sheet1 = $book.create_worksheet :name => '问卷答题'
$sheet1.column(0).width = 70
$sheet1.column(1).width = 10
$sheet1.update_row 0, '问题','分数'


format = Spreadsheet::Format.new :color => :black,
                                 :weight => :bold,
                                 :size => 14

$sheet1.row(0).height = 18
$sheet1.row(0).default_format = format

$questions = qs.readlines

class QuestionWindow < FXMainWindow

  def initialize(app)

    super(app, 'Question', :opts => DECOR_ALL, :x => 100, :y => 100)

    build_layout

  end

  def build_layout
    controls = FXVerticalFrame.new(self,
                                   LAYOUT_SIDE_RIGHT|LAYOUT_FILL_Y|PACK_UNIFORM_WIDTH)

    contents = FXHorizontalFrame.new(self,
                                     LAYOUT_SIDE_LEFT|FRAME_NONE|LAYOUT_FILL_X|LAYOUT_FILL_Y|PACK_UNIFORM_WIDTH,
                                     :padding => 20)

    @@counter = -1

    @button = FXButton.new(contents,
                           "开始答题 \n\n\n  (点击按钮进入下一题)",
                           nil,
                           :opts => FRAME_RAISED|FRAME_THICK|LAYOUT_LEFT|LAYOUT_CENTER_X|LAYOUT_CENTER_Y|LAYOUT_FIX_WIDTH|LAYOUT_FIX_HEIGHT,
                           :width => 550, :height => 200)
    @button.connect(SEL_COMMAND) {
      if @@counter == -1 then
        next_question
      end
    }

    qb = FXButton.new(controls, '退出', :opts => FRAME_RAISED|FRAME_THICK|LAYOUT_FILL_X)
    qb.connect(SEL_COMMAND) {
      $book.write '问卷.xls'
      getApp().exit(0)
    }

    group1 = FXGroupBox.new(controls, '选择答案',
                            GROUPBOX_TITLE_CENTER|FRAME_RIDGE)
    @group1_dt = FXDataTarget.new(2)


    b0 = FXButton.new(group1, '0: 该描述完全不符合我的情况', :opts => FRAME_RAISED|FRAME_THICK|LAYOUT_NORMAL|JUSTIFY_LEFT)
    b0.connect(SEL_COMMAND) {
      choice(0)
    }

    b1 = FXButton.new(group1, '1: 该描述基本不符合我的情况', :opts => FRAME_RAISED|FRAME_THICK|LAYOUT_NORMAL|JUSTIFY_LEFT)
    b1.connect(SEL_COMMAND) {
      choice(1)
    }

    b2 = FXButton.new(group1, '2: 该描述有些符合我的情况', :opts => FRAME_RAISED|FRAME_THICK|LAYOUT_NORMAL|JUSTIFY_LEFT)
    b2.connect(SEL_COMMAND) {
      choice(2)
    }

    b3 = FXButton.new(group1, '3: 该描述符合我的情况', :opts => FRAME_RAISED|FRAME_THICK|LAYOUT_NORMAL|JUSTIFY_LEFT)
    b3.connect(SEL_COMMAND) {
      choice(3)
    }

    b4 = FXButton.new(group1, '4: 该描述非常符合我的情况', :opts => FRAME_RAISED|FRAME_THICK|LAYOUT_NORMAL|JUSTIFY_LEFT)
    b4.connect(SEL_COMMAND) {
      choice(4)
    }
  end

  def choice(x)
    return if @@counter == -1
    $sheet1.update_row @@counter+1, "#{@@counter+1}: #{$questions[@@counter]}" , x
        next_question
  end

  def next_question
    @@counter = @@counter+1
    if (@@counter < $questions.length) then
      @button.text = "#{@@counter+1} / #{$questions.length}:\n #{$questions[@@counter]}"
    else
      puts '问卷答题写入 excel 文件: 问卷.xls'
      $book.write '问卷.xls'
      getApp().exit(0)
    end
  end

  def create
    super
    show(PLACEMENT_SCREEN)
  end

end

if __FILE__ == $0
  application = FXApp.new('Question', 'Question')

  QuestionWindow.new(application)

  application.create

  application.run
end
