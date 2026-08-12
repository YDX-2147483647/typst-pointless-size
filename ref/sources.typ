#let pt = "pt"
// 描述点数时，优先使用原文说法，若不针对特定原文，则写成 $1 pt$。不要写成 #1pt，因为这种写法会缩减精度，例如 #7.875pt 会被显示成 7.88pt。

// 某些资料自身前后不一致，或者情况比较复杂。这时可在 notes 中利用以下`draw`函数描述情况
#import "visualize.typ": draw-as-log-period
#let draw(csv, width: 20em, height: 12em) = figure(draw-as-log-period(
  std.csv(bytes(csv)).map(((g, p)) => (g, float(p))),
  width: width,
  height: height,
))


/// 数据源
#let sources = (
  基准: (
    brief: [CTeX 宏包既有规则],
    notes: [
      === 与此无矛盾的实作

      - *MS Word*#footnote[仅限中文版 MS Word。按 @张小衡2006 的说法，英文版 MS Word 不支持字号，并且如果先用中文版设置号数，再用英文版打开，那么设置中会显示点数。另外，当前微软的 PowerPoint、Excel、画图均无号数可选，苹果 iCloud 上的 Pages 也无号数可选。]

        #link("https://support.microsoft.com/zh-cn/office/更改字号-931e064e-f99f-4ba4-a1bf-8047a35552be#:~:text=以 .5 的倍数表示（例如 10.5 或 105.5）")[MS Word 的点数只能精确到 $0.5 pt$]，#link("https://practicaltypography.com/point-size.html#:~:text=Word%20lets%20you%20spec%C2%ADify%20point%20sizes%20in%20half%2Dpoint%20in%C2%ADcre%C2%ADments%2E")[英文版也如此]。输入 10.4 会报错「编号无效」。

        - 某杂志编辑部 @刘岱伟2001 整理了 MS Word 97。与基准相比，除了缺少八号，其余数值、定义范围完全相同。缺少八号应该只是省略了#footnote[这篇文章转录@source:朱永和1999 就省略了小七、特号等。]，因为很多资料都显示 MS Word 97 有八号，比如下图#footnote[拍摄于我校图书馆开架阅览TP317.2。]联想集团总策划、总代理《无老师电脑自学丛书——文字处理Word》（丛书主编吴文虎，教育科学出版社，1997年10月第1版，1997年10月第1次印刷，ISBN 7-5041-1753-6）57页的图7.8。

          #figure(image("assets/Word-联想1997-八号.png", width: 80%))

        - 香港某大学 @张小衡2006 测试了 Windows 上的 MS Word 2003，发现数值、定义范围与基准完全相同#footnote[不过该文行文很不靠谱，还说 PowerPoint、Excel 也适用，]。

        - 我自己检查了 Windows 上的 MS Word 2019，发现数值、定义范围#footnote[小三比较特殊。菜单上有「小三」，但单击「增大字号」和「减小字号」按钮会跳过小三。增减字号按钮生成的不变子序列是「……、3、4、八号、七号、小六、六号、8、小五、10、五号、11、小四、四号、三号、小二、20、二号、小一、一号、28、小初、48、72、……」。]与基准完全相同。检查数值的方法是创建空文档，填充任意文本并设置字号，然后运行以下脚本。

        #figure({
          set text(0.6em)
          show: columns.with(2)
          ```python
          """扫描当前目录`a.docx`中的字号"""
          # /// script
          # requires-python = ">=3.14"
          # dependencies = ["pywin32>=311"]
          # ///
          from pathlib import Path
          import win32com.client

          def get_all_font_sizes(docx_path):
              word = win32com.client.Dispatch("Word.Application")
              word.Visible = False

              try:
                  doc = word.Documents.Open(docx_path)
                  whole_doc = doc.Range()  # 获取整个文档的范围

                  sizes = set()
                  start_pos = 0
          ```

          colbreak()

          ```python
                  while start_pos < whole_doc.End:  # 遍历所有字符范围
                      # 创建一个小范围
                      char_range = doc.Range(start_pos, start_pos + 1)
                      if char_range.Font.Size:
                          sizes.add(char_range.Font.Size)
                      start_pos += 1

                  doc.Close(SaveChanges=False)
                  return sizes

              except Exception as e:
                  print(f"错误: {e}")
                  return []
              finally:
                  word.Quit()

          sizes = get_all_font_sizes(str(Path("a.docx").resolve()))
          print(sizes)
          ```
        })

      - #link(
          "https://www.wps.cn/learning/question/detail/id/2940",
        )[*WPS* 学堂 → 技巧问答 → 其他 → 操作编辑，2020-05-07]

        从此文计算公式看，点制是 $1/72 "in"$。

        不过 WPS 学堂的说法未必符合 WPS 实际情况，参考@source:石家庄2001 中的对比。需要再测试一下。

      - *LibreOffice*

        #link("https://github.com/LibreOffice/core/blob/e0bb9dbaea421053e70600c2ac8737fcadf1ed58/svtools/source/control/ctrltool.cxx#L777-L801")[源代码`ctrltool.cxx`]如下，其中定义了```cpp struct ImplFSNameItem```，点数精确到 $0.1 pt$，号数采用 UTF-8 存储；```cpp aImplSimplifiedChinese```记录了号数与点数的映射关系，数值、定义范围与基准完全相同。根据 git 记录，#link("https://github.com/LibreOffice/core/commit/14e52d3b28336f693b35ef50d4938d2602581744")[这一功能添加于2000年12月7日（UTC+8）]，最初有繁简两版，二者没有实质性区别，只是「号」「號」写法不同。

        #figure({
          set text(0.6em)
          show: columns.with(2)
          ```cpp
          struct ImplFSNameItem
          {
              sal_Int32   mnSize;
              const char* mszUtf8Name;
          };

          const ImplFSNameItem aImplSimplifiedChinese[] =
          {
              {  50, "\xe5\x85\xab\xe5\x8f\xb7" },
              {  55, "\xe4\xb8\x83\xe5\x8f\xb7" },
              {  65, "\xe5\xb0\x8f\xe5\x85\xad" },
              {  75, "\xe5\x85\xad\xe5\x8f\xb7" },
          ```
          colbreak()
          ```cpp
              {  90, "\xe5\xb0\x8f\xe4\xba\x94" },
              { 105, "\xe4\xba\x94\xe5\x8f\xb7" },
              { 120, "\xe5\xb0\x8f\xe5\x9b\x9b" },
              { 140, "\xe5\x9b\x9b\xe5\x8f\xb7" },
              { 150, "\xe5\xb0\x8f\xe4\xb8\x89" },
              { 160, "\xe4\xb8\x89\xe5\x8f\xb7" },
              { 180, "\xe5\xb0\x8f\xe4\xba\x8c" },
              { 220, "\xe4\xba\x8c\xe5\x8f\xb7" },
              { 240, "\xe5\xb0\x8f\xe4\xb8\x80" },
              { 260, "\xe4\xb8\x80\xe5\x8f\xb7" },
              { 360, "\xe5\xb0\x8f\xe5\x88\x9d" },
              { 420, "\xe5\x88\x9d\xe5\x8f\xb7" }
          };
          ```
        })

      - #link("https://typst.app/universe/package/pointless-size")[*Typst* pointless-size 默认规则]

        当初开发时，已经注意到号数并无统一标准，所以专门参考了 CTeX、MS Word、WPS 学堂等。

      === 关于 CTeX

      #link("https://web.archive.org/web/20080705111113/http://svn.ctex.org/aloft/Packages/ctex/trunk/ctex.dtx")[2008年7月5日（UTC）网页存档中 svn 仓库的`ctex.dtx`]#footnote[该网页会乱码。可先按 Latin-1 编码，再按 GB 18030 解码，恢复出正常版本。]、#link("https://github.com/CTeX-org/ctex-kit/blob/063f230872de2acb08243f8f488b2ec32860861f/ctex/ctex.sty#L140-L155")[2009年5月5日（UTC）git 仓库的`ctex.sty`]、#link("https://svn.tug.org:8369/texlive/trunk/Master/texmf-dist/tex/latex/ctex/def/ctex-common.def?pathrev=14341&view=markup#l71")[2009年7月20日（UTC）`ctex`宏包在CTAN发布时的`ctex-common.def`]，全部采用这套规则。这些文件中的规则采用 $1/72 "in"$ 点，但专门转换成了 TeX `pt`表达。根据网页存档`ctex.dtx`中以下这段注释，很可能最初2004年2月13日采用 TeX `bp`（即 $1/72 "in"$），后来2004年5月13日才改成 TeX `pt`。

      #figure(
        ```latex
        % \changes{v0.2b}{2004/02/13}{修改缺省的行距}
        % \changes{v0.4}{2004/05/13}{中文字号定义改为直接使用~pt~为单位}
        ```,
      )

      另外，#link("https://svn.tug.org:8369/texlive/trunk/Master/texmf-dist/doc/latex/ctex/ctex.pdf?revision=14341&pathrev=14341&view=co")[2009年7月20日（UTC）CTAN上2009年6月26日v0.92版《`ctex`宏包说明》]有以下两段，似乎2004年4月23日设置`c5size`时五号是 $10 pt$ 而非 $10.5 pt$？为简单，此处忽略这种可能。

      #quote(attribution: [2 使用帮助 → 2.3 选项 → 2.3.1 只能用于文档类的选项])[
        下面的选项可能会是最经常使用的。但是它们只能用于文档类（`ctexart`、`ctexrep`和`ctexbook`）。

        / cs4size: 使用小四字号为缺省字体大小。
        / c5size: 使用五号字为缺省字体大小。#text(font: "KaiTi")[这个是`ctex`文档类的缺省模式。]
      ]
      #quote(attribution: [3 版本更新，节选与字号相关的项目])[
        - v0.4 2004/05/13 General: 如果指定了标准的 LaTeX 字体大小，则不使用中文字号；中文字号定义改为直接使用`pt`为单位；`\zihao`: 删除`\CTEX@fontsize`命令，改为直接使用`\fontsize`命令

        - v0.2d 2004/04/23 General: Change option c5size to base on 10pt basic class；补上字号定义中行间距参数中缺少的`\CTEX@bp`；修改缺省的字号大小

        - v0.2b 2004/02/13 General: 修改缺省的行距；修改缺省的字号大小
      ]

      此外，按照以上《宏包说明》「3 版本更新」最后一条 v0.0 2003/04/26，这一套 CTeX 代码始于2003年。不过根据 @farn1990 文末参考文献列表和 #link("https://www.gbv.de/dms/tib-ub-hannover/238526763.pdf")[1988 International Conference on Computer Processing of Chinese and Oriental Languages, August 29 - September 1, 1988 Toronto, Canada 之目录]以下这段，这次会议 Session IIh Input/Output 来自台北的 Tzao-Lin Lee 做了题为 CTeX: A TeX-based Word Processing System 的短报告。二者相差十五年，大概率没有关联吧。

      #figure(grid(
        columns: 2,
        image("assets/Toronto-1988-CTeX-节选.pdf", page: 1), image("assets/Toronto-1988-CTeX-节选.pdf", page: 2),
      ))
    ],
  ),
  CTeX新: (
    brief: [CTeX 宏包新增加的实验性规则],
    notes: [
      #set par(justify: false) // font-size-system 那段太难对齐了

      #link("https://github.com/CTeX-org/ctex-kit/issues/543")[ctexsize: 重设各级字号大小 · ctex-kit\#543] 决定增加`experiment/font-size-system=traditional`，随后 #link("https://github.com/CTeX-org/ctex-kit/issues/813")[ctex-kit\#813] 从`traditional`改成了`letterpress`。
    ],
  ),
  新CCT: (
    brief: [新版CCT所用规则],
    via: [#link("https://mirrors.cernet.edu.cn/ctex/cct/cct-0.618033-3-win32.zip")[从校园网联合镜像站下载`/ctex/cct/cct-0.618033-3-win32.zip`]，解压，查看`files/tex/latex/cct/CCT.cfg`],
    notes: [
      CCT 是#link("https://zh.wikipedia.org/wiki/LaTeX#CCT")[「最早支持简体中文的TeX」，「由中国科学院数学与系统科学研究院的张林波研究员编写」]，#link("https://liam.page/2013/10/15/LaTeX-CCT-template/")[在2013年左右就已经过时]。根据#link("http://maths.nju.edu.cn/~meijq/tex/NewCCTreadme.pdf")[张林波《关于新版CCT的说明》（2006年2月3日）]，CCT有新老两个版本，老版最终版本号为 5.14，新版版本号类似 0.6180-3。《说明》还指出，新版CCT支持以下两种排版流程。

      - 第一种继承自老版，基于CCT自己的预处理程序`cct`，导出CCT专用格式的dvi文件。

        这种方式调用CCT字库，字号设置源于`cct.dat`，其格式类似CSV/TSV。

      - 第二种是新版新增，用TeX宏文件`CCT.sty`替代原有预处理程序，导出兼容于CJK宏包系统的普通dvi文件。

        这种方式调用CJK宏包系统的字库，字号设置取决于`CCT.cfg`调用的`\CCTdefzihao`命令。

      此处记录第二种，@source:老CCT 记录第一种。

      根据镜像站文件系统数据和 #link("https://ctex.org/ctex/release-notes/")[CTeX 套装更新记录]，CCT 0.618033-3 发布于2024年5月；不过根据`CCT.cfg`中的注释，新版CCT的字号表从2005年8月开始就没改过了。

      以下是`CCT.cfg`中的相关内容。

      #figure(
        ```latex
        % $Id: CCT.cfg,v 1.4 2005/08/31 14:33:16 zlb Exp $
        %
        % Configuration file for CCT.sty >= 0.6.0
        %
        % Note: \CCTdefzihao and \CCTdefziti only apply to CJK fonts.
        %

        % CCT字号表: \CCTdefzihao 字号 磅数

        \CCTdefzihao	0	36	% 初号
        \CCTdefzihao	1	27	% 一号
        \CCTdefzihao	2	21	% 二号
        \CCTdefzihao	3	16	% 三号
        \CCTdefzihao	4	13.75	% 四号
        \CCTdefzihao	-4	12	% 小四号
        \CCTdefzihao	5	10.5	% 五号
        \CCTdefzihao	-5	9	% 小五号
        \CCTdefzihao	6	8	% 六号
        \CCTdefzihao	7	5.25	% 七号
        ```,
      )
    ],
  ),
  老CCT: (
    brief: [老版CCT所用规则],
    via: [同@source:新CCT，但查看`files/tex/latex/cct/cct.dat-dist`],
    notes: [
      CCT有新老两个版本，此处是老版。详见@source:新CCT。

      与@source:新CCT 相比，老CCT的字号整体小一圈。不过新老CCT调用的字库不同，不排除印出来差不多大。

      `cct.dat`文件内容如下。

      #figure(
        ```txt
        10
        8.5     8.5      0.06             0.2            -5
        11.32   11.32    0.06             0.2            -4
        34.0    34.0     0.06             0.2             0
        26.0    26.0     0.06             0.2             1
        20.0    20.0     0.06             0.2             2
        15.0    15.0     0.06             0.2             3
        13.0    13.0     0.06             0.2             4
        9.9     9.9      0.06             0.2             5
        7.5     7.5      0.06             0.2             6
        5.0     5.0      0.06             0.2             7
        -----------------------------------------------------------------
        width   height   h_space_factor  v_space_factor  zihao

          space=width*h_space_factor,
          stretch=sapce*0.5,
          shrink=space*0.3333333333,
          depth=height*v_space_factor

        $Id: cct.dat,v 1.2 2004/11/26 11:41:32 zlb Exp $
        ```,
      )

      以上文件显示最后修改于2004年11月26日。根据#link("http://maths.nju.edu.cn/~meijq/tex/CCTmanual.pdf")[张林波《科技排版软件TeX中文接口——CCT DOS 版参考手册》（1997年11月）]中展示的`CCT.DAT`，当时CCT的宽高点数和号数定义范围就是以上这样了，只不过`-4`、`-5`两行分别排在`4`、`5`之后（而非排在`0`之前），同时行距与字高之比（`v_space_factor`）是`0.0`（而非`0.2`）。

      《手册》还指出，此处字宽、字高的单位是 TeX `pt` $1/72.27 "in"$。

      另外，根据以下邓建松、彭冉冉，陈长松《LaTeX2e科技排版指南》（科学出版社，2001年9月第一版，2001年9月第一次印刷，ISBN 7-03-009239-2）#footnote[拍摄于我校图书馆开架阅览TS8。]，`cct.dat`定义的点数并不绝对，还会按`\magnification`参数统一放缩。

      #figure(grid(
        columns: 2,
        ..range(4).map(n => image("assets/LaTeX_CCT-邓建松2001.pdf", page: n + 1))
      ))
    ],
  ),
  天元: (
    brief: [天元（TY）2.62 版所用规则],
    via: [前往#link("https://web.archive.org/web/20210418112817/http://wims.math.ecnu.edu.cn/ty/tydownload.php")[「天元（TY）下载之页」网页存档]，下载「TYWIN (TY for Windows) 2.62 版」中的#link("https://web.archive.org/web/20210418112817/http://wims.math.ecnu.edu.cn/ty/tywin263.zip")[`tywin263.zip`]，解压，查看其中的`TYWIN.HLP`文件#footnote[#link("https://en.wikipedia.org/wiki/WinHelp")[WinHelp文件`*.hlp`]是微软的一种专有格式帮助文件，相关支持已于2015年Windows 10 移除。可以通过#link("https://www.herdsoft.com/ftp/downloads.html#hlp2rtf")[`hlp2rtflx-2.16.tar.gz`等软件]转为RTF查看。]],
    notes: [
      天元是和CCT同时代的中文TeX系统，同样早已过时。#link("https://ask.latexstudio.net/ask/article/92.html")[网传天元初始作者肖刚已于2014年病逝。]

      `TYWIN.HLP`结尾标注「（2003年4月29日更新）」，不过在陈志杰、赵书钦、万福永《LaTeX 入门与提高》（高等教育出版社，2002年2月第1版，2002年2月第1次印刷，ISBN 7-04-010468-7）#footnote[
        这是此书第一版，以天元为主，顺带介绍了CJK宏包与CCT。后来出了第二版（高等教育出版社，2006年5月第2版，2006年5月第1次印刷，ISBN 7-04-019379-5），题名不变，作者增多为陈志杰、赵书钦、李树钧、万福永，内容改为以CJK宏包为主，不再介绍天元与CCT。CJK宏包不支持号数制，作者在第二版30页「§3.2 中文字体」中表示：「中文书籍基本字号是五号字，大小为 $10.5$ 磅（$1$ 磅即`1bp`）， 很接近于西文基准尺寸之一的 $11 pt$（严格讲是 $10.9 pt$），即使不再另行定义汉字字号的尺寸，选择 $11 pt$ 的基准尺寸，排版输出的结果也是令人满意的。如果一定要全部使用各个汉字字号规定的尺寸，则需定义一系列的命令。」「现有的最大字体尺寸命令是`\Huge`, 略小于一号汉字（$28$ 磅）。如果需要更大的汉字，例如初号汉字（$42$ 磅），就需要自行定义一个字体尺寸命令。`\fontsize{字体尺寸}{行距}`」
      ]中，已有数值、定义范围相同的字号规则。

      以下是`TYWIN.HLP`相关内容节选。（标点符号码位和颜色等格式有改动）
      #set raw(lang: "latex")

      #quote(attribution: [概述])[
        1999年6月……为了使用LaTeX 用户的方便，把汉字的尺寸作了改变，使得在`magnification=1000`时默认的汉字大小（对应于`\标`）为高`ht=8.07pt`，深`dp=1.807pt`，宽`wd=9.877pt`，相当于五号字的大小。而TYWIN 在`magnification=1095` 时默认的汉字大小为高`ht=8.191truept`，深`dp=1.927truept`，宽`wd=10.118truept`。为了区分，把修改后的软件称为“天元排版”（TYPB）。它与TYWIN 的唯一区别就是汉字大小不同。TYWIN适合在`magnification=1095`下使用，而TYPB 则适合在`magnification=1000`时使用。相应地TYWIN 使用的汉字表格文件是`TYCFNT.TAB`，而TYPB 使用的表格文件名为`TYPBFNT.TAB`。

        2000年5月和8月又对TYWIN作了两次修改，使得用户可以自定义字体命令以及汉字大小。这样就使得TYPB失去了存在的必要。详情可参见配置节。新的版本也取消了对输出密度的限制，因此你只需把密度设为1016DPI，依次用天元和TeX处理后，把生成的DVI文件用DVIPS输出成PS文件，就可以在任何方正系统的照排机上出片，达到专业级的印刷质量。本人主编的《高等代数与解析几何》就是这样做的……
      ]
      #quote(attribution: [配置])[
        *汉字大小：*本页供有特殊需要的用户使用。可以设定汉字的字宽、字深、字间距以及瘦体系数及阔体系数。汉字总是方的，字宽减去字深就是字高。框内数字置`0`意为取默认值。每次修改后所有的汉字会重新生成。当`magnification=1095`时默认值对应于原TYWIN的设定，当`magnification=1000`时默认值对应于原来的TYPB的设定。
      ]
      #quote(attribution: [天元中的汉字尺寸])[
        天元中确定汉字尺寸的命令有：

        ```
        \半(=\五): 5 pt  \六: 6 pt  \七: 7 pt   \八(=\小): 8 pt
        \九: 9 pt        \标(默认): 10 pt       \中: 12 pt
        \大: 14 pt       \特: 17 pt \双: 20 pt  \巨: 25 pt
        \叁: 30 pt       \肆: 40 pt \伍: 50 pt  \陆: 60 pt
        ```

        #set par(justify: false)
        在`\magnification=1095`，或在LaTeX中设定`\documentclass[10pt]{article}`或`\documentstyle[10pt]{article}`的条件下，有以下的对应关系（第三列是相应的LaTeX命令，AMSTeX中不能用）：

        ```
        初号     \肆                    一号     \巨      \Huge
        二号     \双      \huge         三号     \特      \LARGE
        四号     \大      \Large        小四号   \中      \large
        五号     \标      \normalsize   小五号   \九      \small
        六号     \八(=\小)\footnotesize 七号     \五      \tiny
                 \七      \scriptsize
        ```
      ]
      #quote(attribution: [天元与CCT的转换])[
        (1) 从天元转换成CCT时……程序按如下对应关系修改：

        ```
        \半\五\六\七  -> \zihao{7}                     \八\小       -> \zihao{6}
        \九          -> \zihao{-5}                    \标          -> \zihao{5}
        \中          -> \zihao{-4}                    \大          -> \zihao{4}
        \特          -> \zihao{3}                     \双          -> \zihao{2}
        \巨          -> \zihao{1}                     \叁\肆\伍\陆 -> \zihao{0}
        ```

        (2) 当把CCT源文件转换成天元源文件时……并按如下对应关系修改

        ```
        \zihao{7}   -> \七                            \zihao{6}   -> \小
        \zihao{-5}  -> \九                            \zihao{5}   -> \标
        \zihao{-4}  -> \中                            \zihao{4}   -> \大
        \zihao{3}   -> \特                            \zihao{2}   -> \双
        \zihao{1}   -> \巨                            \zihao{0}   -> \叁
        ```
      ]
    ],
  ),
  方正书版: (
    brief: [《方正书版9.1实用教程》],
    via: link(
      "https://github.com/CTeX-org/ctex-kit/issues/543#issuecomment-2848708469",
    )[「宁波，晓舟」提供照片，Explorer-cc 转发],
    notes: [
      有可能是高萍2003年（中国环境科学出版社，ISBN 7-80163-577-9）版本。

      #figure({
        set image(width: 60%)
        image("assets/方正书版-a.png")
        image("assets/方正书版-b.png")
      })

      另外，从#link("https://www.ecsponline.com/goods.php?id=95577")[科学出版社]可以下载到#link("https://www.ecsponline.com/yz/BCCB7AB3899504A75868D141B7EB4E1A8000.pdf")[高萍等《方正排版标准教程·方正书版9.11》的样章]（2003年10月第一版第一次印刷，ISBN 7-03-012034-5），其10–11页（PDF 26–27页）「1.4.5 字号的选用」与以上照片的内容几乎完全相同，只有表格编号、Point/point大小写不同，如下。

      #figure(grid(
        columns: 2,
        image("assets/方正书版2003-a.png"), image("assets/方正书版2003-b.png"),
      ))

      需要注意，以上恐怕都不是方正官方资料。#link("https://www.founderfx.cn/resources/support/方正书版2008-210使用指南.pdf")[方正官方《方正书版2008使用指南》（2011年10月）]似乎没有这种映射表格，`HT`（汉体）等注解（类似 LaTeX 控制命令，例如`HT5SS`指定5号书宋）中的字号有常用字号（只写数字）、磅字号（数字后加`.`）、级字号（数字后加`j`）三种填写格式；不过这份资料可能不是详细说明书，比如其中有`HT4"W`这种用法，但似乎并未解释 `4"` 代表小四。
    ],
  ),
  朱永和1999: (
    brief: [表1，#link("https://doi.org/10.16811/j.cnki.1001-4314.1999.02.004")[朱永和《排字的点数制与号数制探讨》]，1999年第2期《编辑学报》],
    via: [@朱永和1999],
    notes: [
      表1「号数制与点数制的换算关系」有号数制、字身尺寸（mm）、点数（P）、点数制尺寸（P）四列，这里只用第一列和第四列。最后两列的区别是精确度不同，第三列精确到千分位（例：18.197），而第四列精确到 0.25（例：18.25）。

      #figure(image("assets/朱永和1999.png", width: 60%))

      #quote(attribution: [§2.5 点数制比号数制更便于字体变形])[
        例如用 10 P 宋体和 10.25 P 仿宋体或 10.5 P 楷体混排时，我们见到的是字身尺寸不一致但视觉效果上大小相一致的不同字体；而 10 P 宋体和 10 P 仿宋体或 10 P 楷体混排时，我们见到的则是字身尺寸一致但视觉效果上大小不完全一致的不同字体。
      ]
    ],
  ),
  方正飞某: (
    brief: [方正飞腾和方正飞翔的官方资料],
    via: [Z-Library 影印PDF和#link("https://www.founderfx.cn/product/1012.jhtml")[方正飞翔产品中心]],
    notes: [
      北大方正电子有限公司#link("https://z-lib.sk/book/M9W6g7ml53/飞腾实例教程.html")[史晓岩、陈明《北大方正飞腾排版系统教学丛书——飞腾实例教程》]（2000年2月第1版第1次印刷，电子工业出版社，ISBN 7-5053-5727-1 / TP·2951）大部分按磅数（甚至精确到百分位），不过19页（PDF 29页）提到1号是27.5磅。
      #figure(grid(
        columns: 2,
        image("assets/飞腾实例教程-前言.png"), image("assets/飞腾实例教程-王选.png"),
        image("assets/飞腾实例教程-19.png"), image("assets/飞腾实例教程-94.png"),
      ))

      不过同系列#link("https://z-lib.sk/book/ezXkmDVR9Z/飞腾标准教程.html")[汤帜、党筱菁、任力《飞腾标准教程》]（2000年1月第1版第1次印刷，电子工业出版社，ISBN 7-5053-5728-X / TP·2952）80–81页（PDF 83–84页）显示5号是10.48磅，3号是15.72磅。这似乎暗示磅数有精度限制，取不到.5或.75。
      #figure(grid(
        columns: 2,
        image("assets/飞腾标准教程-80.png"), image("assets/飞腾标准教程-81.png"),
      ))

      #link("https://founderfx-storage.oss-cn-zhangjiakou.aliyuncs.com/shuomingshu/方正飞翔9.0印刷版说明书.pdf")[《方正飞翔9.0印刷版使用说明书》（2024年4月）]122页（PDF 140页）字体字号命令一节展示了五号是10.5磅。

      #figure(image("assets/方正飞翔.png", width: 80%))
    ],
  ),
  方正跨媒介: (
    brief: [15页（PDF 23页）表1-2-1「常用号数对应的磅数」，#link("https://www.tup.tsinghua.edu.cn/upload/books/yz/095514-01.pdf")[杨雷鸣、贾皓、梅林、李谦《方正飞翔跨媒介出版实用教程》]，清华大学出版社，2022年12月第1版，2022年12月第1次印刷，ISBN 978-7-302-61595-8],
    via: [#link("https://www.tup.tsinghua.edu.cn/booksCenter/book_09551401.html")[清华大学出版社图书详情] → 资源下载 → 样章下载；国家图书馆总馆北馆开架阅览TS8亦有],
    notes: [
      此书介绍方正飞翔，由北京北大方正电子有限公司组织编写，网传杨雷鸣曾任方正开发部部长。

      此书虽然是官方资料，但未必靠谱。以下正文写「常用的号数」有九种，而表1-2-1也是「常用号数」，却不止九种。正文还写Word也如此称谓，但所给数值、定义范围与Word（@source:基准）均不完全一致，至少多了「特号」「大一号」等。

      另外，此书所写一号 $28 pt$ 与@source:方正飞某\并不相同。

      #quote[
        号数制：以铅活字的大小用号来称谓的体制。常用的号数有九种：一号至七号，还有小五号和小四号。除了排版软件以外，在Word软件中，字号也是用号数制称谓的。

        在排版过程中，偶尔也会涉及号数制与点数制的换算，这里提供常用号数对应的磅数，以便可以快速查询，如表1-2-1所示。
      ]

      #figure(grid(
        columns: 2,
        ..range(4).map(n => image("assets/方正跨媒介-节选.pdf", page: n + 1))
      ))
    ],
  ),
  石家庄2001: (
    brief: [表一，#link("https://doi.org/10.3969/j.issn.1000-663X.2001.11.017")[刘韬、师彦茹《方正书版应用杂谈》]，2001年第11期《中国印刷》],
    via: [@刘韬2001],
    notes: [
      作者单位是石家庄陆军学院印刷厂。

      表一是「常用级数、号数、磅数、字号对照表」，文中说明如下。这里只用号数、磅数两列。

      #quote[
        其中“级数”是指手动照排的字号及尺寸，“号数”是指铅字的字号及尺寸，“字号”是指北大方正激光照排系统的字号及尺寸，“磅数”是指由铅字尺寸折算的磅数。
      ]
      #quote[
        特别是有些用户，因受铅活字影响较大，目前不但坚持着铅活字的一些提法，甚至还总是觉得自动照排的字号不“标准”。
      ]
      #quote[
        通过表一可以看出，虽然广泛应用于各印刷厂的方正激光照排系统的字号，与相应的铅活字字号和手动照排字号的尺寸都不相同，但方正书版系统是可以模拟铅活字和手动照排字号尺寸的。
      ]
      #figure(image("assets/石家庄2001.png", width: 80%))

      #set math.frac(style: "horizontal")
      注意表中磅数和「约合毫米」的比值并不一致，例如 $9.665 "mm" / 28.50"P" = 0.339 "mm/P"$，但 $14.761 "mm" / 42.0"P" = 0.351 "mm/P"$。不清楚作者是怎么算的。

      另外，「约合毫米」为 $3.698$ 这行的「级数」「号数」「磅数」「字号」均为空，而 $3.25$ 这行同时有「13J」和「5号」。按常理推断，应该是「5行」从 $3.698$ 行错误地串到了 $3.25$ 行。以下胡宏芳《方正书版排版技术基础与实例教程》（电子工业出版社，2004年11月第1次印刷，ISBN 7-121-00457-7）#footnote[拍摄于我校图书馆开架阅览TS8。]123页表5-2「常用的汉字字号」介绍了方正书版的「字号」「注解写法」「毫米（mm）」「点阵数（字身）」对应关系，其中注解`5`也对应 $3.698 "mm"$。另外与以上相比，胡宏芳书还多了 $63$、$72$、$84$、$96$。

      #figure(grid(
        columns: 2,
        ..range(4).map(n => image("assets/方正书版-飞思2004.pdf", page: n + 1))
      ))

      #quote[
        我们以笔划最饱满的黑体字为例，通过仔细的对比发现，WPS2000系统的字号与方正书版系统的相应字号，无论提法或是尺寸基本上都是一致的。只是WPS2000系统字体的笔划更细，在视觉上显得字心略大一些。Word97系统的字号与方正书版系统的字号不太一致，通过我们对样张的测量发现，Word97的“初号”字为 $14.595 "mm"$、“一号”字为 $9.36 "mm"$、“小三号”字为 $5.15 "mm"$，它的“小初号”字与方正书版系统的“0号”字相同、“八号”字与方正书版系统的“7号”字相同，其他字号与相应的方正书版系统字号相同。也同WPS2000一样，因其字体的笔划略细，在视觉上字心显得更大一些。
      ]

      此外，@刘岱伟2001 也整理了北大方正号数与点数的映射关系，除了六号对应7.75而非7.85，其余数值全部与以上一致，定义范围也相同。
    ],
  ),
  政府: (
    brief: [国家标准 #link("https://std.samr.gov.cn/gb/search/gbDetailed?id=BBE32B661B7E8FC8E05397BE0A0AB906")[GB 40070—2021《儿童青少年学习用品近视防控卫生要求》]和行业推荐性标准 #link("https://std.samr.gov.cn/hb/search/stdHBDetailed?id=8B1827F23645BB19E05397BE0A0AB44A")[CY/T 154—2017《中文出版物夹用英文的编辑规范》]],
    via: [#link("https://std.samr.gov.cn")[全国标准信息公共服务平台]可看],
    notes: [
      #quote(attribution: [CY/T 154—2017])[
        11.2.1 中文文本中夹用英文时，英文字号应与中文字号匹配。常用的为：中文“小五号”与英文“9P”相对应，中文“五号”与“10.5P”相对应。
      ]
      #quote(attribution: [GB 40070—2021])[
        *4.3 正文汉字、字母和阿拉伯数字用字*

        4.3.1 小学一、二年级用字应不小于16P（3号）字……注：P——Point，$1 "P" approx 0.35 "mm"$。

        4.3.2 小学三、四年级用字应不小于14P（4号）字……

        4.3.3 五～九年级和高中用字应不小于12P（小4号）字……

        *4.5 目录、注释和拼音等辅文用字*

        ……小学阶段最小用字应不小于10.5P（5号）字，初中和高中阶段最小用字应不小于9P（小5号）字。

        *8 学习用报纸卫生要求*

        8.2 ……用字字号应不小于10.5P（5号）字。

        *9 学龄前儿童学习读物卫生要求*

        9.3 字号应不小于16P（3号）字……辅文用字应不小于10.5P（5号）字。
      ]

      此外，国家推荐性标准 #link("https://std.samr.gov.cn/gb/search/gbDetailed?id=71F772D7E77CD3A7E05397BE0A0AB82A")[GB/T 9704—2012《党政机关公文格式》]也间接描述了号数与点数的映射关系：

      #quote[
        *5.2.1 页边与版心尺寸*

        ……版心尺寸为 $156 "mm" times 225 "mm"$。

        *5.2.2 字体和字号*

        ……一般用3号仿宋体字。特定情况可以作适当调整。

        *5.2.3 行数和字数*

        一般……每行排28个字，并撑满版心。
      ]

      #set math.frac(style: "horizontal")
      这样算下来三号字的宽度不超过 $156 "mm" / 28 = 15.793 pt$（按 $1 pt = 1/72 "in"$ 计算）。

      又，#link("https://std.samr.gov.cn/gb/search/gbDetailed?id=71F772D78446D3A7E05397BE0A0AB82A&review=true")[GB/T 12200.2—1994《汉语信息处理词汇　02部分：汉语和汉字》]「4 术语和定义 → 4.1 汉语和汉字 → 4.1.2 文字 → 4.1.2.9 字号 character number」展示了字号大小，不过未定义点数。手工测量像素数的话，比例大约是 $36.1 : 32.0 : 28.1 : 21.2 : 16.1 : 14.2 : 12.4 : 10.5 : 9.28$。若假设五号是 $10.5 pt$，则数值比较接近@source:方正书版（特别是初号、小初），但定义范围只有初号到六号与小初、小四、小五。

      #figure({
        image("assets/GB_T 12200.2—1994-a.png", width: 80%)
        image("assets/GB_T 12200.2—1994-b.png", width: 80%)
      })
    ],
  ),
  CLReq-main: (
    brief: [#link("https://www.w3.org/TR/clreq/#considerations_in_designing_type_area")[§7.1.1.5 基本版式设计的注意事项 - 中文排版需求 | W3C 小组备忘草稿]，2026-05-03 版],
    notes: [
      #quote[“号”由于当年金属活字各地厂家的规范不一而不尽相同……不作为规范性规定。]

      部分字号在表中给了两种数值#footnote[
        原文写法形如 27.5/28 pt，这里将斜线理解成「或」而非除号。
      ]，此处记录点数更接近整数的那种，而 @source:CLReq-extra 记录另一种。注意 CLReq 只是写一号、二号、三号、四号、六号各有两种数值，并未暗示这两种数值是否各自配套。

      CLReq 并未写明出处，不过曾在 #link("https://github.com/w3c/clreq/issues/142")[Type size equivalents between 号 and "point" · Issue #142 · w3c/clreq] 讨论过，其中有以下这段。

      #quote[
        In hot metal era, the type size are also different from different type foundries. See the values shown in the book below, it just showed one of the variations. Some of the values go with Word, some of them not.
        #figure(image("assets/clreq-issue-142.png", width: 4em))
      ]
    ],
  ),
  CLReq-extra: (
    brief: [CLReq 表格另一种],
    notes: [
      详见 @source:CLReq-main。
    ],
  ),
  Ken-2-JP: (
    brief: [480页 Table 7-3. The G typographic unit 的 Japan 部分，Ken Lude _CJKV Information Processing_，2008年12月第二版，O'Reilly Media，ISBN 978-0-596-51447-1],
    via: [ArchiBC 买过这本书，提供了2009年的PDF],
    notes: [
      那张表的 China 部分和@source:基准\完全相同，无定义的情况也相同。

      #figure(image("assets/Ken 2.png"))

      479页起介绍号数制：
      #quote[
        #set text(lang: "en")
        The other typographic unit, which I am calling G, is unique in that its scale is the reverse of what one finds in the other typographic units. That is, the higher the value of G, the smaller the size… G is used strictly for type size. This typographic unit, as used in Japan, was developed by Shozo Motoki (本木昌造 motoki shōzō) in 1933, and was supposedly based on a traditional Japanese metric system called Kujira-Jaku (鯨尺 kujira jaku). The Chinese developed a comparable system with the same name, but it differs in that fractional, referred to as "small" (小 in Chinese), values are included within its scale. For example, in Chinese, the 0G value is referred to as 初号, and the fractional or "small" version is referred to as 小初.
      ]

      其中1933这个年份恐怕有误，因为#link("https://ja.wikipedia.org/wiki/本木昌造")[本木昌造]1875年就去世了。

      这本书有 #link("https://www.oreilly.com/catalog/errata.csp?isbn=9780596514471")[Errata]，不过除了以上这个年份错误，confirmed 和 unconfirmed 都没提到480页或相关内容。

      这本书相关章节还被#link("https://ccjktype.fonts.adobe.com/2009/04/post_1.html")[字体度量单位 - CJK Type Blog | Adobe]（2009年4月2日发布）引用了。该网页也展示了表格，不过格式不同，并且中国版少了小一（日本版本来就没有小一）。

      #figure(image("assets/CJK Type Blog.png", width: 80%))
    ],
  ),
  Ken-1: (
    brief: [342页 Table 7-2. The G Typographic Unit，@source:Ken-2-JP 那本书的第一版（推测是1999年那版，但未验证）],
    via: [r 买过纸书，提供了照片],
    notes: [
      与第二版 @source:Ken-2-JP 相比，此处第一版没有区分中国和日本，并且7G、8G的位置、数值不同。

      #figure(image("assets/Ken 1.jpg", width: 80%))
    ],
  ),
  姜别利: (
    brief: [十九世纪西人设计重要中文字体号数对照表，#link("https://www.thetype.com/2016/12/11232/")[罗佳洋《从「拼合」到「格致」：有关西人汉字认知的设计史叙述——以「拼音」「拼合字」为例》第三章]，2016年12月7日],
    notes: [
      另有 $36 pt$ Two-line Great Primer《华英字典》大号字体（马礼逊，1815年）和 $9 pt$ Bourgeois（柯尔，1851年）。

      从西文名可看出#footnote[其实 Brevier = $8 pt$ 和 Pica = $12 pt$，详见 #link("https://en.wikipedia.org/wiki/Traditional_point-size_names")[Traditional point-size names - Wikipedia]。]：
      - Two-line Pica#footnote[@source:jawiki-旧\说一号是 Two line English，即 English（四号）的两倍，与此不矛盾。]（一号）略大于 Two-line Small Pica（二号）
      - Two-line Small Pica（二号）是 Small Pica（五号）的两倍
      - Two-line Brevier（三号）是 Brevier（六号）的两倍

      #figure(image("assets/姜别利.png", width: 80%))

      #quote[ 「号数制」是由姜别利自 1858 年起根据英美点制陆续建立起来的]

      #quote[
        就此，特将 19 世纪西人设计汉字大小情况扼要如下（按，附有号数者即属姜别利之号数制[注]）

        [注]：
        #quote[
          此表参冯锦荣先生研究中征引的 1865 年版姜氏《花图（样本）书》史料。（冯锦荣：《姜别利（William Gamble, 1830–1886）与上海美华书馆》，复旦大学历史系、出版博物馆编：《历史上的中国出版与东亚文化环流》，上海：百家出版社，2009 年，第 309–310 页）；并参何步云：《中国活字小史》#footnote[即@source:小史1981 第300页的表1。二者所记一号至七号的点数一致，但@source:小史1981 没有另外记 $36 pt$ 和 $9 pt$。]，上海新四军历史研究会印刷印钞分会编：《活字印刷源流》，北京：印刷工业出版社，1990 年，第 76 页；Walter Henry Medhurst, China, Its State and Prospects, with Especial Reference to the Coast of China in 1831, 1832 & 1833, London: John Snow, 1838, pp556-557；Samuel Wells Williams, "Movable Types for Printing Chinese", Chinese Recorder, VI, 1875.
        ]
      ]

      另外，按照#link("https://www.thetype.com/typechat/ep-135/")[《字谈字畅 \#135：显明解行号号珍》]的说法，姜别利那时点的概念还比较模糊，选择字号并不是按点数选的。不过根据以下西野嘉章编《歴史の文字　記載·活字·活版》（東京大学総合研究博物館）#link("https://web.archive.org/web/20070901124726/https://www.um.u-tokyo.ac.jp/publish_db/1996Moji/05/5901.html")[第三部「活版の世界」]表4印刷品实测结果，以上点数还是比较靠谱的。

      #quote[
        #set text(lang: "ja")
        表4中＊印は一九〇九年十二月二五日刊行の『活版術』（韓国龍山印刷局、隆煕三年刊）に収録されているデータで、印刷局の技手高木徳太郎がこの年の一月に東京の印刷局でポイント計器を使って測定したものであると書かれている（ただし四号のポイント換算は一四ポイント台と誤記）。この測定結果は矢野道也著『印刷術上巻』（丸善、一九二五年刊）に誤記のまま掲載されているが、こちらは東京の印刷局の小山初太郎の測定となっている。近年では矢作勝美著『明朝活字』（平凡社、一九七六年刊）にも誤記のままやはり使われているが、その出典は『印刷局研究所調査報告』第一号（明治四二年九月）所載小山初太郎の「活字型の説話」とある。

        #figure(
          caption: [表4　伝道会印刷所の印刷物での実測値と日本国内の活字メーカーの活字の実測値比較],
          image("assets/歴史の文字 - 活版の世界 - 表4.jpg"),
        )

        同＊＊印は矢野道也著『印刷術發達史』（大阪出版社、一九二七年刊）五七～六〇頁に掲載されている、東京大阪の著名活字製造所五杜（社名は伏せてある）の活字をマイクロメーターで測定したものである。原表では活字の大きさと幅を別々に表記しているが、本表では両者の上限と下限を載せた。

        同＊＊＊印は『活字書体』（株式会社モトヤ、一九六七年刊）所載の「活字寸法表」。

        美華書館の号数系列には倍数関係はなく、かろうじて五号の倍である二号が近似値となる。二号はダブル・スモール・パイカ二二ポイント、五号はスモール・パイカ一一ポイントと名称の上では倍数関係が成り立つが、実寸法では成立しない。

        表4を見てわかることは、三号、四号、五号の三サイズは導入した美華書館のサイズをほぼ保って今日に到っていることである。そしてこの三号、四号、五号をもとにして、後に倍数関係が成り立つように一号、二号を直していったと考えるのが自然である。それを指示したのは本木昌造ではないであろう。本木は『新街私塾餘談』や『新聞雑誌』に美華書館の活字を複製した活字見本を出した明治五年には、崎陽新塾活字製造所の経営一切を平野富二に託し、新街私塾の運営に専念しているのである。
      ]

      以上段落的背景及大意：

      + 关于日本活字号数制的来源，传统说法由三谷幸吉于昭和八年（1933年）提出，认为是本木昌造根据日本布用尺度「鲸尺」算出来的。《歴史の文字》作者发现这种说法与数据不符，认为日本号数制其实沿用自美华书馆的字号。
      + 作者提出的证据之一是印刷品实测结果，即表4「传道会印刷所印刷品的实测值与日本国内活字制造商的活字实测值比较」。表4中括号前是毫米数，括号内是从毫米数换算出的美式点数。表4综合自多篇文献，不同文献的记录格式不同，有些数据还存在误记，作者对此做了些说明。
      + 根据表4，作者发现美华书馆的号数系列没有倍数关系。二号 Double Small Pica 与五号 Small Pica 在名称上是二倍，但实测下来仅勉强近似成立。
      + 从表4还可看出，日本导入了美华书馆号数制，三号、四号、五号基本保持不变。然后以三号、四号、五号为基础，对一号、二号进行修改，使倍数关系成立，「这样想比较自然」。做出这一指示的应该不是本木昌造。本木在《新街私塾闲谈》和《新闻杂志》上刊登了美华书馆活字的印刷样本，明治五年（1872年），他将崎阳新塾活字制造所的一切经营委托给平野富二，专注于新街私塾的运营。

      《歴史の文字》后文总结：
      #quote[
        #set text(lang: "ja")
        日本の近代活字史は本木昌造を神格化することによって事実から遠くかけ離れたものとなり、後世の研究者も実証的に検証する方法をとらず先人の引用に終始したため、誠に不可思議な説が定説として堂々と罷り通っているのである。
      ]
      中文翻译：
      #quote[
        日本的近代活字史由于将本木昌造神格化而远离事实，后世的研究者也不采用实证性的验证方法而始终引用前人，所以不可思议之说作为定论堂而皇之地流传开来。
      ]

      #set math.frac(style: "horizontal")
      #link("https://zh.wikipedia.org/w/index.php?title=字号_(印刷)&oldid=92592966#起源")[起源 - 字号 (印刷) - 中文维基百科，2026-05-09版]也列出了美华书馆的字号、点数，与以上数值十分接近，但一号 $28 pt$ 除外。一号 $28 pt$ 大概率不对。如下图，#link("https://z-lib.sk/book/V05ew11AzW/铸以代刻傳教士與中文印刷變局.html")[蘇精《鑄以代刻：傳教士與中文印刷變局》]560–561页（PDF 576–577页）图12-1、图12-2给出了美华书馆的活字广告与样本，可见一号前七字「我父在天者願爾」与二号#footnote[
        指和一号同一页的「另有二号」。这个「另有二号」似乎和下一页「第二号」不是同一套字，比如「我」字第四笔提的角度不同。同作者#link("https://z-lib.sk/book/6qBYkkjazJ/美华书馆-档案如是说.html")[《美华书馆：档案如是说》]「3 从档案谈姜别利活字的四个问题 → 二、姜别利制造几种中文活字？→ 图3-1」与这两张图十分类似，其图注写「美华书馆出售活字广告（“另有第二号”为原柏林活字，“第二号”为姜别利改善者）［《中国教会新报》第1年第16期（1868年12月19日）］BFMPC/MCR/CH, 200/8/126, W. Gamble to W. Lowrie, Changhai, February 24, 1868.」，看来确实不同。
      ]前八字「我父在天願爾名聖」在竖直方向几乎等长。这说明一号与二号的比例约为 $8/7$，与《歴史の文字》表4测量出的 $8.55 "mm" / 7.6 "mm" approx 7.88 / 7$ 几乎相等，更接近以上所记 $24 pt / 22 pt approx 7.6 / 7$ 而非中文维基百科所记 $28 pt / 21 pt approx 9.3 / 7$。

      #figure(grid(
        columns: 2,
        image("assets/鑄以代刻-560.png"), image("assets/鑄以代刻-561.png"),
      ))

      此外，美华书馆最初没有 $8 pt$ 活字，所以直接把 $5.5 pt$（Small Ruby 或 Ruby）叫六号。姜别利离开后#footnote[从这个角度说，此数据源代号「姜别利」并不准确。]，美华书馆新造 $8 pt$（Brevier）活字，鸠占鹊巢，$5.5 pt$ 才改叫七号。这导致《鑄以代刻》、@孙明远2024 等文献只提一号至六号，并且若给出点数则一号至五号的点数与以上一致，而六号点数等于以上七号点数。
    ],
  ),
  周承民1988: (
    brief: [表1-2　几种常用字号，#link("https://z-lib.sk/book/3zje0vQX9j/活字排版工艺.html")[周承民等《活字排版工艺——凸制专业》]，1988年2月第一版第一次印刷，印刷工业出版社],
    via: [Z-Library 影印PDF],
    notes: [
      题名与@source:曹洪奎1979 相同，内容也高度接近，不清楚是什么原因。

      36–41页#footnote[41页第二行引用了图1-21，但图1-21位于49页，内容是「铅（水）线的种类和形状」。应该是引串了。]（PDF 48–53页）讲「第一章　基础知识 → 第三节　活字 → 三、活字的规格 →（二）活字的大小」。表1-2有名称、字号、尺寸（毫米）、折合点数四列，并注「在括号内的数字是国家出版局办法的活字标准」。此处记录的是字号、折合点数两列，其中折合点数优先采用括号内有效数字位数更多的版本。

      #quote(attribution: [内容提要])[
        本书是由文化部批准，文化部出版事业管理局组织编写的印刷技工学校专业教材之一，供印刷技工学校和印刷职工业余教育排版专业试用，也适用于活字排版工人、技术人员阅读。
      ]

      #quote(attribution: [37页])[
        我国原有七种字号：一、二、三、四、五、六、七号（因七号字太小，只有排在五号字的公式中，用七号来做角标字，也用六点字来代替使用，六点字主要是排版权之用)。后来根据书刊美化版面的需要，又增添一些字号，如八号、小五号、小四号、小二号、小初号、初号、特号和特大号等。
      ]

      #let img(n, ..args) = image("assets/周承民1988-{}.png".replace("{}", n), ..args)

      #layout(size => figure({
        grid(
          columns: (1fr,) * 2,
          align: top + center,
          grid.cell(colspan: 2, img("a", width: 50%)),
          img("b"), img("c"),
          img("d"), img("f"),
          grid.cell(colspan: 2, rotate(90deg, img("e", height: size.width * 80%), reflow: true)),
        )
      }))

      另外，#link("https://github.com/CTeX-org/ctex-kit/issues/543")[tanukihee 曾提供「1993年」周承民《活字排版工艺》的照片]，不过#link("https://github.com/CTeX-org/ctex-kit/issues/813#issuecomment-4412583072")[后来沟通发现很可能不是1993年版]。具体而言，#link("http://opac.nlc.cn/F/RU3P8SNJEGUCQDY58XIF2P2MUK36IGNT5QSCCPGILGVLUNLL22-04993?func=short")[在国家图书馆馆藏目录只能查到@source:周承民1988、@source:曹洪奎1979 两版《活字排版工艺》，而无1993年任何版本]，而且 tanukihee 当时所发两组照片也与这两版分别一致。
    ],
  ),
  曹洪奎1979: (
    brief: [曹洪奎《活字排版工艺》，轻工业出版社出版，张家口地区印刷厂印制，封面写1979年，内页写1983年9月第一版第三次印刷#footnote[根据前言，该书原名《铅印排版技术》，后来利用原纸型重印。因此1979年可能是原版年份，而1983年可能是重印年份。]，统一书号15042·1509],
    via: [#link(
        "https://ss.zhizhen.com/detail_38502727e7500f2685813c708ce0786ae71c99137e1f57f01921b0a3ea25510134114c969f2eae5c409d083e1d75cb511ee4bf8e7dfbe7254ed60b8f61ffee76030c1ddc408afcba31a0e7397ecd1407?&apistrclassfy=0_18_19",
      )[汇雅电子书影印PDF]#footnote[该地址是我校图书馆统一检索页面，不确定校外能否访问。]],
    notes: [
      题名与@source:周承民1988 相同，内容也高度接近，不清楚是什么原因。

      这本书36–39页讲「第二章　排字基础知识 → 三、字号」。

      36–37页正文提及了一号至八号，初号、特号、特大号，小五号、小四号、小二号、小初号，五行、六行、七行，不过只明确指出了「小□号」和「□行」的点数。

      37页表2-3「几种常用字号」有名称、字号、字面、字坯、折合毫米数、折合点（磅）数、等于水线（根）七列，其中字号一列的「小□号」全标了“`*`”，下注「`*`为点数制的铅字。每点为0.35毫米，见第四章外文排字常识。」，而折合点数一列部分数值标有括号而无解释。此处记录的是字号、折合点数两列，其中折合点数优先采用括号内有效数字位数更多的版本。表2-3给出的点数与前文无矛盾，不过表2-3跳过了小二号。

      此节没有明确给出八号对应的点数，不过38页指出一根水线是「五号字的 $frac(1, 8, style: "skewed")$」，且「八号字是3根线稍强」。

      39页表2-4「各号字的相互比例参考表」的注指出：「此表是字面大小比例，在计算版面行长时便于参考，字身的大小稍有误差。」

      #figure(grid(
        columns: 2,
        image("assets/曹洪奎1979-封面.png"), image("assets/曹洪奎1979-内容提要.png"),
        image("assets/曹洪奎1979-前言.png"),
        ..range(36, 40).map(n => image("assets/曹洪奎1979-{}.png".replace("{}", str(n)))),
        image("assets/曹洪奎1979-120.png"),
        image("assets/曹洪奎1979-121.png"),
      ))
    ],
  ),
  叶重光1996: (
    brief: [87页（PDF 101页）表5「铅字规格大小对照表」，#link("https://z-lib.sk/book/r9bD6YveqB/印刷出版插图与版式设计.html")[叶重光、叶朝阳《印刷出版插图与版式设计》]，1996年6月第一版第一次印刷，印刷工业出版社，ISBN 7-80000-207-1 / TS·146],
    via: [Z-Library 影印PDF],
    notes: [
      这本书大量描述了字号之间的倍数关系，提到新1号、新2号而未明确给出点数。

      #quote(attribution: [69–70页（重新分过段，但未改动文字及标点符号）])[
        #set math.frac(style: "horizontal")
        汉字字号的名称。

        汉字活字是以铅铸成的，铅字的名称按大小分为初号、小初号、1号、2号、3号、4号、小4号、5号、小5号、6号、7号……这传统铅铸汉字的大小字号之间，有个特殊的倍数关系。

        - 如将全部标准汉字用缩去 $1/4$ 保留 $3/4$ 的办法，2号字可缩为标准3号字，3号字可缩为标准新4号字，5号字可缩为标准6号字。用缩去 $1/4$ 的办法，可将现成的标准汉字缩出三个比原字小的标准小号汉字；

        - 如将标准汉字用缩去 $1/3$ 保留 $2/3$ 的办法，头号字#footnote[原文如此。从上下文推断，「头号」大概指1号。]可缩为标准新2号字，新1号字可缩为标准的3号字，3号字可缩为标准的5号字，4号字可缩为标准的新5号字，新4号字可缩为标准的6号字，6号字可缩为标准的7号字。用缩去 $1/3$ 的办法，可将现成的标准汉字缩出六个比原字小的标准小号汉字；

        - 如将标准汉字用缩去 $1/2$ 的办法，初号字可缩为标准的新2号字#footnote[原文如此。从上下文推断，此处「标准的新2号字」应是「标准的2号字」。]，头号字可缩为标准的4号字，新1号字可缩为标准的新4号字，2号字可缩为标准的5号字，新2号字可缩为标准的新5号字，3号字可缩为标准的6号字，5号字可缩为标准的7号字。用缩去1/2的办法，可将现成的标准汉字缩出七个比原字小的标准小号汉字；

        - 如将标准汉字用缩去 $2/3$，保留 $1/3$ 的办法，头号字可缩为标准的新5号字，新1号字可缩为标准的6号字，3号字可缩为标准的7号字。用缩去 $2/3$ 的办法，可将现成的标准汉字缩出三个比原字小的标准小号汉字。

        这四个缩比当中，以缩去 $1/3$，保留 $2/3$ 和缩去 $1/2$ 的用途最广，其中以缩去 $1/2$ 的用途最大，这七个级差完全够有级差的制图用。

        汉字的这个特殊倍数关系使出版制图放大缩小有了科学依据，即把难画的小图放大画，贴上同倍的大字再同时缩小刚刚达到图进版心，字为标准字号，整齐标准的目的，实现制图标准化。
      ]

      #figure({
        grid(
          columns: 4,
          ..range(7).map(n => image("assets/叶重光1996-节选.pdf", page: n + 1))
        )
        grid(
          columns: 2,
          ..range(7, 17).map(n => image("assets/叶重光1996-节选.pdf", page: n + 1))
        )
      })

      @刘岱伟2001 表1「文献[1]」如下这列转录了这本书，但将6号 7.875 点误作 7.75，并且文末参考文献列表将出版社误记为北京工业出版社。@刘岱伟2001 表1其它列也都略有出入：「北大方正」列与@source:石家庄2001 中的「号数」「磅数」两列相同，但六号原文 7.85 被记成 7.75；「Word97」列同@source:基准\但缺少八号；「文献[6]」列同@source:朱永和1999 但缺少小七、特号等。

      #figure(image("assets/刘岱伟2001.png"))
    ],
  ),
  小史1981: (
    brief: [309页表2我国现用活字大小称谓、种数及其点数规格一览表，300–312页（PDF 354–366页）何步云《中国活字小史》，#link("https://z-lib.sk/book/Gz3M7GWm5E/中国印刷年鉴-1981.html")[《中国印刷年鉴1981》]],
    via: [Z-Library 影印PDF],
    notes: [
      按照 @高洋2025 的引用，《中国活字小史》也收录于上海新四军历史研究会印刷印钞分会《活字印刷源流》（印刷工业出版社，1990）。不清楚两版是否有区别。

      表2记录了正方、狭长、扁宽三种字身形式的活字，这里只记录正方#footnote[狭长、扁宽的号数覆盖范围小于正方，而且纵向点数与正方一致，唯一例外是狭长小初号有 $31.5 times 24$ 与 $30 times 24$ 两种规格，前者的纵向点数与正方小初号 $30$ 不同。]。另外，表2中初号与二号之间是「大号」，但正文似无此说法，所以这里将大号当作一号理解（仅限大号、小大号；特大号不当作「特一号」理解）。

      #quote(attribution: [304页])[
        现代活字版印刷术是帝国主义者作为文化侵略工具而传入我国的，帝国主义侵略我国，是军舰、大炮和圣经，两手同时使用的。圣经——大炮——圣经，就是侵略我国的公式之一。活字版印刷术就成为印圣经和进行文化侵略的必要工具。
      ]
      #quote(attribution: [304–305页])[
        1845年，花华圣经书房迁宁波，改名美华书馆。于1859年（清咸丰九年）迁至上海……当时陆续制成大小活字七种、每种各定一中国名称。……后来为了称呼方便，从大字到小字就以一号至七号次序，作为我国活字称谓，一直沿用到现在。这七种字，规格都比现用的七种活字高些、大些（见表1）。

        …………

        1869年（清同治八年）6 月，甘勃尔离职返美时，路经日本长崎#footnote[关于「路经」说法，@高洋2025 考证发现姜别利于1869年10月底从美华书馆辞职，1869年秋（十一月）东渡日本，1870年春末（五月）回到上海，1870年八月初启程返美，所以在日本停留了半年，并非单纯「路经」。]，将电镀制模、铸造活字技术及七种铸好活字、传给日本活版传习所的木本昌造。

        这时，日本正逢维新，印刷工业及知识界认为中国的字体不够齐整美观，重新刻制。并改变七种活字字身的大小高低规格，就成为现在中日两国通用的七种汉文活字。……“五·四”后，又向日本购入新四号、新五号、新二号等几种点数制字模。从此活字规格就成为点数、号数混合使用，增加了排版的不便并积压材料、资金。
      ]
      #quote(attribution: [308页])[
        我国活字大小分档，当初只有一号到七号七种（实际上七号字极少用）。"五·四"以后，从日本引进9点、12点、18点、27点、36点等点数系列字模五种，另加狭长体七种，扁体一种，到解放时大小档数共有二十种（均以有字模的计算，只有活字的不计）。

        现用活字大小档数有：正方字身的十八种，狭长字身的十一种（同号数两种规格的有三种，三种规格的有一种），扁宽字身的十种（同号数两种规格的有一种），大小规格共有四十五种（见表2）。
      ]


      #figure(grid(
        columns: 2,
        ..range(13).map(n => image("assets/中国印刷年鉴 1981 - 中国活字小史 - 何步云.pdf", page: n + 1))
      ))
    ],
  ),
  王选1982: (
    brief: [北京大学、潍坊电子计算机厂等于约1982年研发完成的计算机-激光汉字编辑排版系统的改进型#footnote[此前1979年有原理性样机。]],
    via: [《计算机学报》1982年12月23日收到、1984年11月发布的 @王选1984],
    notes: [
      #figure(grid(
        columns: 2,
        image("assets/王选1984-节选.pdf", page: 1), image("assets/王选1984-节选.pdf", page: 2),
      ))

      此文明确指出比例关系按照字心计算。按照 @林川1991#footnote[该文恐怕非常不靠谱。文中表一「对比4#super[[7]]」一列转录了《中国活字小史》（@source:小史1981），但将五号10.5点误作11.5，将特号45点、特初号48点、特中号56点乱作小特42,45、特48、特大56，且无任何说明。在此文末，参考文献列表还将文献 [7] 题名误作「汉字活字小史」，英文介绍更是出现了 theorelical、printting、calaulating 等多处拼写错误与 These foze the writer considers、by way of futes calculation 等莫名其妙的表达。]的说法，字身的比例关系略有不同。

      @林川1991 的依据是#link("https://ss.zhizhen.com/detail_38502727e7500f2685813c708ce0786aa70a95c0efb4cab51921b0a3ea25510134114c969f2eae5cee61fba22a1d40b6bc023b74f5f06b6013b4623dba6cf47e255ac44419becabe9e23de8d5bbbe850?&apistrclassfy=0_18_17")[郭平欣、张淞芝《汉字信息处理技术》]（国防工业出版社，1985年12月第一版，1985年12月第一次印刷，统一书号#footnote[这个「统一书号」似乎是「全国统一书号」，并非ISBN。]15034·2973）中由王选、陈堃𨱇执笔的第十三章「精密汉字编辑排版系统」。此书节选如下，存在以下三处谈及字号。

      - 433页表13-1「印刷用汉字字号与字身点阵的关系」的「字号」「磅数」与王选《计算机学报》文章上的数值、定义范围一致，但「字身点阵大小」比《计算机学报》上的「字心点阵」系统性地大一圈。此外前一页说该表是「按 $29.2 thick frac("线", "毫米", style: "skewed")$ 计算」得出的，似乎暗示还有其它可能。
      - 443页表13-6「字号及其比例关系」与《计算机学报》上的表格雷同，只不过五号一行上方加了条横线。
      - 460页例1提供了初号、头号与二、三、四、小四、五、小五、六、七号的示例，并说系统还允许特大号、特号、小特号、小初号、小二号、小六号。若将「头号」理解成一号，则与《计算机学报》上的定义范围一致。

      #figure(grid(
        columns: 2,
        ..range(7).map(n => image("assets/汉字信息处理技术-1985-节选.pdf", page: n + 1)),
      ))

      王选《计算机学报》文章还引用了欧洲专利 #link("https://worldwide.espacenet.com/patent/search?q=pn%3DEP0095536A1")[EP0095536A1 The representation of character images in a compact form for computer storage]（1982年6月1日申请、1983年12月7日公告，登记号 82302816.2），王选是其唯一发明人。这份专利第2页有 character size（按点数、毫米#footnote[原文为 $"mm"^2$，应该是写错了。]数）和 dot matrix 的对应表格。专利表格中的 character size in point 和《计算机学报》上的「磅数」不全相同（初 35、特 49 在专利表格分别变成了 36、48），dot matrix 与《汉字信息处理技术》中的「字身点阵大小」一致（磅数不同的也一致）。

      #figure(grid(
        columns: 3,
        ..range(3).map(n => image("assets/EP_0095536_A1-节选.pdf", page: n + 1)),
      ))

      《计算机学报》上表格的内容被很多教材直接或间接抄录了，例如：

      - 85页「第3章　汉字字形存储与压缩技术 → 3. 汉字字形的变倍方法 → 表3-19　印刷用汉字的字号表」，李宝安、李燕、孟庆昌《中文信息处理技术——原理与应用》，清华大学出版社，2005年7月第1版，2005年7月第1次印刷，ISBN 7-302-11200-2 / TP·7394

        该书将小二一行「五号的1.708倍」中的 $1.708$ 误作 $7.708$。

      - 141页「第7章　汉字字形和字形库管理技术 → 7.4.1 汉字的字号 → 表7.4　印刷用汉字的字号表」，朱巧明、李培峰、吴娴《中文信息处理技术教程》，清华大学出版社，2005年9月第1版，2005年9月第1次印刷，ISBN 7-302-11761-6 / TP·7655

        该书给表格每行上下都加了网格线，还在表格下方抄录了《计算机学报》文中「表中的小六号、小二号、小初号和小特号是现在铅字所没有的」这一与@source:沪一厂1978 矛盾的说法。
    ],
  ),
  zhwiki: (
    brief: [#link("https://zh.wikipedia.org/w/index.php?title=字号_(印刷)&oldid=92592966#相关换算")[相关换算 - 字号 (印刷) - 中文维基百科，2026-05-09版]以及更早版本],
    notes: [
      此词条经过多人编辑，目前比较混乱，只好忽略了。历史大概如下。（日期按UTC+8）

      #let link-diff(query, body) = link("https://zh.wikipedia.org/w/index.php?title=點_(印刷)&" + query, body)

      + #link-diff("diff=7314482&oldid=7170428")[2008年6月20日Voidvector]向今日名为「點（印刷）」的词条添加了「中国相对」一节，其唯一内容是包含「印刷字号」「排版磅数」两列的表格。该表内容如下，有几点比较特殊：小五是 $9.13$ 而非 $9$，六号 $7.87$ 与小六 $7.78$ 十分接近，无小三、有小六、无八号。
        #draw(
          ```csv
          七号,5.25
          小六,7.78
          六号,7.87
          小五,9.13
          五号,10.5
          小四,12
          四号,13.75
          三号,15.75
          小二,18
          二号,21
          小一,24
          一号,27.5
          小初,36
          初号,42
          ```.text,
        )

      + #link-diff("diff=9494225&oldid=8837213")[2009年3月5日小wing]修改了三号、四号、小五、六号、小六的数值，添加了小三；同时向表格添加了「毫米数」一列，不过七号的毫米数留空了。修改后内容如下。
        #draw(
          ```csv
          七号,5.25
          小六,6.5
          六号,7.5
          小五,9
          五号,10.5
          小四,12
          四号,14
          小三,15
          三号,16
          小二,18
          二号,21
          小一,24
          一号,27.5
          小初,36
          初号,42
          ```.text,
        )


      + #link-diff("diff=10766663&oldid=9494225")[2009年8月2日Lupeng1917]添加了特大号、大特号、特号、八号，并补充了七号的毫米数。修改后内容如下。
        #draw(
          ```csv
          八号,4.5
          七号,5.25
          小六,6.5
          六号,7.5
          小五,9
          五号,10.5
          小四,12
          四号,14
          小三,15
          三号,16
          小二,18
          二号,21
          小一,24
          一号,27.5
          小初,36
          初号,42
          特号,54
          大特号,63
          特大号,72
          ```.text,
          width: 26em,
        )

      + #link-diff("diff=28417336&oldid=25433395")[2013年8月30日119.161.132.13]将小四的毫米数从 $4.32$ 修改为 $4.23$，但「排版磅数」保持 $12$ 不变。

      + #link-diff("diff=31578373&oldid=28417336")[2014年6月13日Pengyanan]移动页面「點 (印刷)」至「字型大小」，随后又移回去。根据编辑摘要，「字型大小」会被繁简转换处理成「字号」，所以做了这番修改。今日「#link("https://zh.wikipedia.org/wiki/字号")[字号]」是个消歧义页，指向商业字号、老字号、字号（印刷）、表字与别号。

      + #link-diff("diff=32801423&oldid=31578373")[2014年9月28日49.196.4.237]向「中国相对」一节添加了首段正文，并引用了 @source:Ken-2-JP 对应的 Adobe CJK Type Blog，尽管词条中的数值、定义范围与Adobe页面中日两列均有差异。

      + #link-diff("diff=41317764&oldid=39615686")[2016年8月31日36.48.110.253]向表格添加了「示例」一列，内容是使用相应字号的示例文本。


      + #link-diff("diff=43787308&oldid=42317616")[2017年3月29日Ryukei]删除了「示例」列，将表头「排版磅数」改成了「点数」，同时在「中国相对」一节添加了以下描述，尽管此时表中七号是 $5.25$，其它字号也不全与 MS Word（@source:基准）相同。

        #quote[
          ……參考下表。请注意，这并不能真实反应传统铅字大小的情况，而是出于软件制约，比如 Microsoft Word 就有「字号必须是 0.5 pt 的倍数」这样的制约，因此 Word 无法将七号字忠实再现为 5.25pt 而只能定为5.5 pt，这样的换算导致与「方正飞腾」等其他排版系统中的「号」大小不尽相同。
        ]

      + #link-diff("diff=45605151&oldid=43788362")[2017年8月10日Ryukei]删除了「点数」列，增加了以下三列，而「毫米数」一列保持不变。

        - 「中国点数」如下。这列与原「点数」基本相同，保留了特大号、大特号、特号#footnote[我于2026年5月私下用邮件询问，Ryukei 表示这段是 Lupeng1917 加的，可以请求来源。]，将一号、四号分别改为了 $27.5\/28$、$14\/13.75$#footnote[原文顺序如此。]，将三号、六号分别改为了 $16\/15.75$、$8\/7.875$，将小六从 $6.5$ 改为了 $6.875$。这列与@source:CLReq-main、@source:CLReq-extra 相比，定义范围多了特大号、大特号、特号、小初、小三、小六、八号，二号是 $21$ 而非 $21\/22$，其余数值、定义范围一致。
        - 「Word点数」与@source:基准\的数值、定义范围完全一致。未定义号数在此列留空。
        - 「日本点数」与@source:Ken-2-JP 的数值、定义范围一致，除了初号是 $45$ 而非 $42$#footnote[我于2026年5月私下用邮件询问，Ryukei 表示手头各种资料显示初号是 $42 pt$，那我估计是当时笔误吧。]。未定义号数在此列填了短横线。
        #draw(
          ```csv
          八号,4.5
          七号,5.25
          小六,6.875
          六号,8
          六号,7.875
          小五,9
          五号,10.5
          小四,12
          四号,14
          四号,13.75
          小三,15
          三号,16
          三号,15.75
          小二,18
          二号,21
          小一,24
          一号,27.5
          一号,28
          小初,36
          初号,42
          特号,54
          大特号,63
          特大号,72
          ```.text,
          width: 26em,
          height: 20em,
        )

      + #link-diff("diff=59481676&oldid=49585017")[2020年5月3日Ryukei]从「#link("https://zh.wikipedia.org/wiki/點_(印刷)")[點（印刷）]」词条拆分出了「#link("https://zh.wikipedia.org/wiki/字级")[字级]」和「#link("https://zh.wikipedia.org/wiki/字号_(印刷)")[字号（印刷）]」。于是该表移动到了「字号（印刷）」的「相关换算」一节，不过内容无变化。
    ],
  ),
  enwiki: (
    brief: [Comparison table 的 Chinese system 一栏，#link("https://en.wikipedia.org/w/index.php?title=Traditional_point-size_names&oldid=1347503780")[Traditional point-size names - 英文Wikipedia，2026-04-07版]],
    notes: [
      此栏引用了@source:CLReq-main，但与@source:CLReq-main 和@source:CLReq-extra 都有差异（比如 CLReq 未定义八号，而这里却定义了），反而与@source:基准\的数值、定义范围完全相同。

      更蹊跷的是此表有 American system、Continental system 和 Chinese system 三栏，把相同点数而非相同尺寸的名称放到同一行。例如初号这一行有42点、"≈ 14.817 mm" (American system)、"≈ 15.8 mm" (Continental system) 几格，只知道初号是42点，而说不清初号具体多大。

      另外，表中点数使用 $7 1/2$ 这种带分数而非 $7.5$ 这种小数，不存在精度问题。

      #figure(image("assets/enwiki - Comparison table - Traditional point-size names.pdf"))

      #link("https://en.wikipedia.org/wiki/East_Asian_typography#Font_sizes")[Font sizes - East Asian typography - Wikipedia] 也列表展示了号数制，数值、定义范围完全相同，不过引用的是 @张小衡2006。
    ],
  ),
  jawiki-旧: (
    brief: [1967年以前，#link("https://ja.wikipedia.org/w/index.php?title=活字&oldid=108907778#号数活字")[号数活字 - 活字の大きさ（活字大小） - 活字 - 日文维基百科，2026-03-26版]],
    notes: [
      分了「旧号数，1967年以前」「新号数，1967年以后」两个表格，不过无引注，文中也没介绍1967年发生了什么。两表只有初号这一行相同，此处记录前者，@source:jawiki-新\记录后者。

      表格注释说是按美式点，不过本文件记录时没有和 $1/72 "in"$ 区分。

      #figure(image("assets/jawiki.png", width: 60%))

      旧号数与@source:Ken-2-JP 的数值、定义范围完全相同。
    ],
  ),
  jawiki-新: (
    brief: [日文维基百科所记1967年以后的版本],
    notes: [
      见@source:jawiki-旧。

      由于新旧点数十分接近，*此处将「新□号」当作「□号」处理*，而未像其它数据源那样当作「小□号」处理。

      此外，#link("https://kikakurui.com/z8/Z8305-1962-01.html")[JIS Z 8305：1962 活字の基準寸法]（活字的标准尺寸）规定了活字的点数、毫米数和允差，如下图。表中点数分两列：左列全部为整数；右列有零有整，并且与日文维基百科所记新点数完全对应（尽管这份标准本身完全没有出现「号」字，更未描述点数与号数如何对应）。另外，表下的注释好像说右列尽量不要使用，但该文件此处图文分割乱了，难以看清。

      #figure(image("assets/JIS Z 8305-1962.png"))
    ],
  ),
  神田: (
    brief: [日本东京神田的株式会社錦精社《各種活字標準規格表》],
    via: [#link("https://github.com/CTeX-org/ctex-kit/issues/543#issuecomment-747950396")[tanukihee 提供照片]，但#link("https://github.com/CTeX-org/ctex-kit/issues/813#issuecomment-4412583072")[具体来源、年代已不清楚]],
    notes: [
      #figure(image("assets/神田.jpg", width: 60%))

      照片中表格和表格下方不完全一致。表格中有新六号 $8 pt$、六号 $7.875 pt$；而表格下方只有6号，并标注是8号的两倍。然而小六应小于六号，再考虑@source:jawiki-旧、@source:jawiki-新\的说法，这里*将原文新六号 $8 pt$ 记作六号，而忽略原文六号 $7.875 pt$*。
    ],
  ),
  东京1942: (
    brief: [17–18页（PDF 37–38页）「把舊有的鉛字，用點數體制精密的計算起來，有如次表（以日本東京鑄地#footnote[原文如此，大约是指築地。]活版所鉛字為標準）……」，#link("https://annas-archive.gl/md5/5e8449c0511568a1edb564020a1bbe7c")[俊山幸男《活版印刷術》，蘇士清译]，國立四川造紙印刷科職業學校出版部，民國三十一年九月（1942年9月）初版，#link("https://openlibrary.org/books/OL62414572M")[OL62414572M]],
    via: [
      #show raw: it => {
        show "_": "_" + sym.zws // Allow line break
        it
      }
      Anna's Archive `upload/wikilib_2025_01/Republican_Era_Books_in_the_National_Library_of_China`影印PDF#footnote[
        此文件目前缺少许多元数据，连标题都没有。此书在Anna's Archive另有元数据完善的影印DjVu，但缺少16、17两页之间的「點數體鉛字大小字樣」和最后的「讀者意見表」。
      ]；#link("https://taiwanebook.ncl.edu.tw/zh-tw/book/NCL-9910010362")[臺灣華文電子書庫NCL-9910010362]亦有彩色影印PDF，#link("https://commons.wikimedia.org/wiki/File:NCL-9910010362_%E6%B4%BB%E7%89%88%E5%8D%B0%E5%88%B7%E8%A1%93.pdf")[Wikimedia Commons同名文件]可下载
    ],
    notes: [
      此书6页（PDF 25页）译者注还提及小五是 $9 pt$，不过并非专门介绍東京，故未抄录。

      此书有四百多页，其中以下几节相关。第一编、第二编的作者分别是是宮崎榮太郎、飯田常治，不是同一人，所以两部分不完全搭配。
      #[
        #show super: it => h(0.2em) + it
        - 第一編 鉛字總論，鉛字製造法及字模製造法 → 第一章 鉛字總論#super[1–26页]
          - 一 鉛字和鉛字各部#super[2–4页]、二 鉛字的大小#super[4–6页]、三 點數體制#super[6–22页]

        - 第二編 採字，排字及自動排字機等 → 第二章 鉛字 → 一 鉛字的大小#super[117–124页]
          - A 漢文鉛字的大小#super[118–120页]、B 西文鉛字的大小#super[120–123页]、C 注音用的七號字#super[123页]、D 鉛字大小的不統一#super[123页]、E 點數體鉛字#super[123–124页]
      ]

      #quote(attribution: [6页（PDF 25页）])[
        即相當小派卡的稱為五號，五號的二倍是二號。相當於英格里斯的稱為四號，四號的二倍是一號。相當於普萊維爾的是六號，六號的二倍是三號。於是由一號至七號制定了七種鉛字。其次為了書籍的標題和廣告的組版的應用，又作了一種二號的一倍的初號鉛字，最小的又作了六號的二分之一的八號鉛字，一共九種。在這九種中間，以五號字為本位鉛字，多用於印書籍雜誌新聞等件，七八號字多用經注說明注音等件。

        外國的鉛字種類很多，名稱也各異，近年已完全統一採用點數體制 point system。所有鉛字都稱為「一點」不復用舊式的名稱了。（此點數體制已經傳到東方各國，日本和中國也很盛行，最流行的小五號字即點數體；但一般人多稱為九磅字，概係九「點」point 讀音之訛傳——譯者）
      ]
      #quote(attribution: [7页（PDF 26页）])[
        原來在歐西各國還沒有實行點數體制的時候，所有鉛字的大小都是鑄造所隨意想出來的，並沒有一定標準的根據。東方各國的鉛字既然都是仿效歐西各國製造的，自然也沒有一定的標準，都是嘰哩磆碌雜亂無章的。

        鉛字的大小不止是兩個鑄造所出來的互有差異，就是同一鑄造所的出品因為年代湮久，也多少有些不同之處不能始終如一。
      ]
      #quote(attribution: [19页（PDF 39页）])[
        ……譬如說一號和二號的差別，我們很精確地知道一號是四號的二倍，二號是五號的二倍，它的差別就是四號是五號差的二倍，是這樣複雜的記憶起來的。至四號與五號間的差別全然曖昧不明，那非用精算器計算不可。再有我們熟悉八號#footnote[原文如此。从上下文推断，此处应为「七號」。]的二倍是五號，五號的二倍是二號的複雜關係，但八號七號六號的大小用眼睛一看就知道，那非是内行人不可。……
      ]
      #quote(attribution: [119页（PDF 139页）])[
        初號和二號，二號和五號的關係我們已經知道了，那麼一號二號三號的各個比例如何，也有曉得的必要，但是正確的像前圖那樣幾倍幾倍的比例是不可能的，這是不便的一點。大致說起來有如以次這樣的舉例計算法，就是三號約為五號的一倍半，六號約為五號四分之三，四號約為五號的一倍又八分之三稍小。為再求詳細一點可參考上圖，但上圖只是表示大小的比例，不是實際的大小，實際的大小如以次的字樣：……
      ]
      #quote(attribution: [121–123页（PDF 141–143页）])[
        以次是西文鉛字大小的名稱，羅列的 (1) (2) (3) 是表示大小的順序，各行的下邊相當於上面鉛字的倍數。

        （一）Excelsior —— Nonpareil —— Pica\
        （二）Brilliant —— Minion —— English\
        （三）Semi-brevier —— Brevier（六號）—— 2 line Brevier（三號）\
        （四）Diamond —— Bourgeois —— Great Primer\
        （五）Pearl —— Long Primer —— Paragon\
        （六）Ruby（七號）—— Small Pica（五號）—— 2 line Small Pica（二號）\

        這和漢文鉛字樣式是同樣的，上下成倍數，至於橫列的 Brilliant, Semi-Brevier, Diamond, 等不成倍數比例。

        ……在前例的西文鉛字名稱行列裏所有括弧裏的倍數是和那種鉛字同大小的漢文鉛字，西文和漢文一起排組的時候必須知道。例如在五號鉛字插入英文的時候以用小派卡鉛字 Small Pica 為最合適。
      ]
      #figure(text(dir: rtl, grid(
        columns: 3,
        // 裁掉一点白边
        inset: -1em,
        ..range(1, 7).map(p => image("assets/活版印刷術1942-节选.pdf", page: p)),
        grid.cell(colspan: 2, image("assets/活版印刷術1942-节选.pdf", page: 7)),
        ..range(8, 15).map(p => image("assets/活版印刷術1942-节选.pdf", page: p)),
      )))
    ],
  ),
  秀英1903: (
    brief: [ZIP 0008–0009页，Price List of Printing Type 歐文角及定價，株式會社秀英舍鑄造部·活版製造所·#link("https://archive.org/details/seibundo1903specimen")[製文堂《活版見本帖》]，明治三十六年二月（1903年2月#footnote[日本从明治六年（1873年）开始，月份与公历完全同步，不再是日月共同影响的阴阳历月份。]）],
    via: [Internet Archive 的#link("https://archive.org/download/seibundo1903specimen/seibundo1903specimen_images.zip/")[`seibundo1903specimen_images.zip`]#footnote[Internet Archive 亦提供PDF，但不如`*_images.zip`清晰。]],
    notes: [
      这是个三百多页的非卖品小册子，设计比较特殊。左侧封面写 Type Specimens，开头是「歐文角及定價」和各种西文字体样张；右侧封面写「活版見本帖」与「和漢字體」，开头是各种汉字与假名字体样张；中间还有围棋盘、装饰线、花形图案等的样张，以及位于ZIP 0106页的版权页。

      原文中七号的点数使用 $5 1/2$ 这种带分数而非 $5.5$ 这种小数，不存在精度问题；另外原文将一号写作「壹號」，但二号至八号的数字并未写成贰叁肆伍陆柒捌。

      #figure({
        grid(
          columns: 5,
          gutter: 0.5em,
          ..(1, 2, 5, 6, 7).map(p => image("assets/seibundo1903spec-节选.pdf", page: p)),
        )
        grid(
          columns: 2,
          ..(3, 4).map(p => image("assets/seibundo1903spec-节选.pdf", page: p))
        )
      })
    ],
  ),
  京新厂1981: (
    brief: [北京新华字模厂《雕刻字字模字体——样本》，1981],
    via: [山东一位老师傅赠与 r，r 扫描了 PDF 并提供#footnote[此PDF从内部资料节选而来，所以存在一些断链。]；扫描时还摆放了黑白双色比例尺#footnote[r在 commit message 中提示：「测量时请优先以色块宽度为准，最好避免使用刻线。」]],
    notes: [
      该文件规格页未写明号数，所以 PDF 提供了样本页，并覆盖了所有存在的字号。不过「我厂常用字模规格」表格与样本页略有出入。规格表称「点数」，而样本页称「○磅」；规格表有48点，而样本页最大是「初号 42磅」；规格无24点，而样本页有「新一号 24磅」。

      样本页有「大一号」「一号」「新一号」三种一号。根据这位师傅的说法，他们一般*把 $30 pt$ 标称「大一号」的字模叫「一号」*，$27.5 pt$ 标称「一号」的字模叫「新一号」，*印刷标准要求使用一号字的地方会使用前者*。此处按文件原文录入，没有考虑这种错位。

      「我厂常用字模规格」表格有点数，字身（边长）、字面（边长），模体（长、宽、厚），模深（总深、座深），头长等列，表头右侧注明「毫米」。规格表的点数、字身边长、字面边长正相关。点数、字身边长近似成正比，系数 $(351.46 plus.minus 0.04) "μm" \/ pt$；但字面与字身边长不完全成正比，从 $5.25 pt$ 到 $48 pt$，字面与字身边长之比从 $87%$ 大致单调递增到 $95%$。

      #figure(grid(
        columns: 2,
        ..range(3, 9).map(p => image("assets/京沪印刷厂.pdf", page: p))
      ))
    ],
  ),
  沪一厂1972: (
    brief: [「各体铅字字型、字数、供应品种参照表」，上海字模一厂（华丰铸字制模厂）《字模与铅字样本》，1972],
    via: [同@source:京新厂1981，r 提供影印 PDF；Anna's Archive `duxiu/initial_release`亦有#link("https://annas-archive.gl/md5/6b6959ea87efa54c653edb08cee4d241")[上海字模一厂《字模与铅字样本》全文影印PDF]],
    notes: [
      与@source:沪一厂1978 是同一厂不同年代的文件。

      该表是产品名录，所以很多点数、号数映射关系描述了多次。正方的字号前后一致；长体、扁体字号的纵向点数前后一致，而横向点数未必一致；正方、长体、扁体都有时，三者纵向点数一致。例如对应24页的四号仿宋是 $14$，对应28页的四号长仿宋是 $14 times 10.5$，二者纵向点数相同；而同样是小特，对应34页的长牟是 $42 times 36$，对应47页的隶书是 $42 times 63$，对应51页的黑变是 $42 times 31.5$，三种纵向点数相同但横向点数不同。另外，对应28页的五号长仿宋，不知为何只给了一个点数，表格和样本页都如此。

      长体、扁体出现了正方没有的字号，包括特中、小特、小初、小大、小二。为尽可能完整，这里不仅记录了正方点数，还记录长体、扁体了独有号数的纵向点数。

      以下是r提供的影印本。

      #figure(image("assets/京沪印刷厂.pdf", page: 9, width: 80%))

      以下是 Anna's Archive 影印本。尚未仔细核对，不过两份影印本的内容很可能完全相同。另外，Anna's Archive 影印本显示此文件有「各种铅字及空衬铅规格表」，包含「点数」「号别、品种」「字身厚度」「字身大小」四列，并且表格下方注明「$1$ 点 $= 0.35$ 毫米」——这与@source:沪一厂1978 很类似。

      #figure(grid(
        columns: 4,
        gutter: 0.5em,
        ..range(12).map(p => image("assets/沪一厂1972 - annas-archive - 节选.pdf", page: p + 1)),
      ))
    ],
  ),
  沪一厂1978: (
    brief: [「铅字品种规格参照表」，上海字模一厂（华丰铸字制模厂）《字模与铅字》，1978],
    via: [同@source:京新厂1981，r 提供影印 PDF],
    notes: [
      与@source:沪一厂1972 是同一厂不同年代的文件，1978年文件只是增加了七号 $6 pt$ 和数学符号 $5.25 pt$（后者未计入 data）。

      该表原文出现了「七行（特大号）」「五行（特号）」「四行（初号）」，这里按括号中的号数录入。

      该表存在和@source:沪一厂1972 一样的重复现象，并且排列方式暗示同一号数正方、长体、扁体的纵向点数相同。该表长体、扁体也出现了正方没有的字号，包括特中号、小特号、小初号、小大号。为尽可能完整，这里不仅记录了正方点数，还记录了长体、扁体独有号数的纵向点数。

      表中「字身厚度」「字身大小」毫米数分别与纵向、横向点数对应，对应关系符合表下所注「$1$ 点 $= 0.35$ 毫米」。

      #figure(grid(
        columns: 2,
        ..range(10, 12).map(p => image("assets/京沪印刷厂.pdf", page: p)),
      ))
    ],
  ),
  丹江厂1980: (
    brief: [「铅字品种规格参照表」，湖北丹江文字六〇五厂（上海字模二厂）《字模与铅字样本》，1980年11月],
    via: [同@source:京新厂1981，r 提供影印 PDF],
    notes: [
      与@source:沪一厂1972 及@source:沪一厂1978 是临近地区不同厂不同年代的文件。该厂旧称上海字模二厂，不过三线建设时期就迁去了湖北，字模实物上的铭文也写丹江而非上海。

      该表号数全部标在括号内，但个别行还在括号外标注了行数，包括「七行（特大号）」「五行（特号）」「四行（初号）」。这里按括号中的号数录入。

      该表存在和@source:沪一厂1978 一样的重复现象和排列方式。该表扁体出现了正方没有的小大号，点数为 $24 times 36$。为尽可能完整，这里不仅记录了正方点数，还记录了扁体独有号数的纵向点数。

      纵向点数为 $30$ 的三行未标注号数，此处未记录。

      表中「字身厚度（毫米）」「字身大小（毫米）」分别与纵向、横向点数对应，系数是 $350 "μm" \/ pt$。

      #figure(grid(
        columns: 2,
        ..range(12, 16).map(p => image("assets/京沪印刷厂.pdf", page: p)),
      ))
    ],
  ),
  沪新厂TTK: (
    brief: [末页表格的「新华厂」列，#link("https://www.thetype.com/social/know-20260116/")[The Type — 中文字号解读·知道Know]#footnote[现在该网页的```html <body>```似乎无法正常加载，但```html <head>```正常。]，2026年1月16日],
    via: [#link("https://mp.weixin.qq.com/s/q9bkC4q-K3JAkUIkiNcDMw")[微信公众号推送《The Type Know · 中文字号解读》]或 #link("https://www.instagram.com/p/DTkErnAEYar/")[Instagram《The Type Know · 中文字号解读》]],
    notes: [
      原文有「头号」而无一号，这里将头号理解为一号。原文还有「6点」，这里未记录。

      #figure(grid(
        columns: 3,
        gutter: 0.5em,
        ..range(6).map(n => image("assets/thetype-know-20260116.pdf", page: n + 1)),
      ))

      定义范围、数值与@source:京新厂1981 并不相同。新华厂在很多地方都有，根据相关情况#footnote[我于2026年5月私下用邮件询问，作者表示：「不同时期的点、号不一致，而同一时期的不同地区也不一致。……对于一号字来说，北京新华厂就是 $27.5 pt$，上海无论是新华厂、商务厂还是中华厂都是 $28 pt$。」]，此处很可能是指上海的。
    ],
  ),
  商务厂TTK: (
    brief: [The Type Know 末页表格的「商务厂」列],
    notes: [
      见@source:沪新厂TTK。原文有「头号」而无一号，这里将头号理解为一号。

      有六号、八号，但无七号，原文如此。
    ],
  ),
  中华厂TTK: (
    brief: [The Type Know 末页表格的「中华厂」列],
    notes: [
      见@source:沪新厂TTK。原文有「头号」而无一号，这里将头号理解为一号。原文还有「6点」，这里未记录。
    ],
  ),
  申报1935: (
    brief: [DjVu 27页本館各種鉛字名稱及面積，#link("https://annas-archive.gl/md5/d60b4e040dc7a3109ffa5983d9bedd9b")[《申報槪况》]，民國廿四年五月（1935年5月）],
    via: [Anna's Archive `upload/duxiu_main`影印DjVu],
    notes: [
      影印本封面「民國廿四年五月」中的「廿四」不太清楚，但根据 DjVu 12页《本館印刷機之進步》最后一栏「去年（民國廿三年）」的说法，封面「廿四」应该确凿无疑。其它各页的照片标题、文章落款、统计图坐标轴等也兼容这一判断。

      原文长度按 $"in"$，这里转换成了 $1/72 "in"$ 点录入。除了五号 $1/7 "in" = (10 + 2/7) times 1/72 "in"$ 以外，其余长度均可无误差表达为 $1/72 "in"$ 点的有限小数倍。

      原文「新二號」前面是「新大號」，但无「大號」「一號」「新一號」，这里将新大號理解为小一。

      #figure(text(dir: rtl, grid(
        columns: 2,
        gutter: 0.5em,
        ..range(4).map(n => image("assets/申报1935-节选.pdf", page: n + 1)),
      )))
    ],
  ),
  报业1936: (
    brief: [#link("https://annas-archive.gl/md5/9dd6afeae7aa5b47ee1c04e88ddf1ed3")[劉覺民《報業管理概論》]，商務印書館，民國二十五年六月（1936年6月）初版，#link("https://openlibrary.org/books/OL62414970M")[OL62414970M]],
    via: [
      #show raw: it => {
        show "_": "_" + sym.zws // Allow line break
        it
      }
      Anna's Archive `upload/wikilib_2025_01/Republican_Era_Books_in_the_National_Library_of_China`影印PDF#footnote[此书在Anna's Archive有多份文件。此文件相对更清晰，尽管目前缺少许多元数据，连标题都没有。]
    ],
    notes: [
      作者刘觉民很可能是参加过辛亥革命的先驱。

      此书多次提及号数与点数的对应关系，但前后并不完全一致。

      第三章提到一号、二号、三号的线度分别是四号、五号、六号的二倍，但第四章正文描述的点数如下图（将「大三十點 30-Point」理解为 $30 pt$），并不符合二倍关系。
      #draw(
        ```csv
        頭號,30
        二號,24
        三號,18
        四號,12
        五號,10
        六號,8
        ```.text,
      )

      第四章表格还介绍了各号铅字每 $"in"$ 能排几字，由此可推出各号数对应的 $1/72 "in"$ 点数。结果如下图，与第四章正文描述略有差异，但也偏离第三章所说二倍关系。
      #draw(
        ```csv
        1号,36
        2号,24
        3号,18
        4号,14.4
        5号,10.285714
        6号,8
        ```.text,
      )

      另外，第三章还描述了甲乙丙丁戊五种新式字模之间的倍数关系。若以戊种的线度作为 $10.5 pt$，则这五种对应的点数如下图。注意此处「新式」并不是指小二、小五。
      #draw(
        ```csv
        甲,42
        乙,31.5
        丙,21
        丁,15.75
        戊,10.5
        ```.text,
      )

      #quote(attribution: [
        第三章　報業組織的實際問題 → 第五節　最低限度的設備\
        70–71页（PDF 77–78页）\
        （重新分过段，但未改动文字及标点符号）
      ])[
        - 銅模的字形大小，舊式的分為一，二，三，四，五，六，等號。一號字一枚大小等於四枚四號字；二號字一枚等於四枚五號字；三號字一枚等於四枚六號字。

        - 近來又有新二號字和新五號字的銅模，比較舊二號舊五號稍小些。

        - 新式的分為甲乙丙丁戊五種。甲種九枚面積等於乙種十六枚，等於丙種三十六枚，等於丁種六十四枚，等於戊種一百四十四枚。換言之：甲種一枚面積合丙種四枚，合戊種十六枚；乙種一枚合丁種四枚，戊種九枚；丙種合戊種四枚。新式的各種大小有一定相互標準較普通的銅模更為悅目。
      ]

      #quote(attribution: [第四章　印刷工場管理 → 第五節　活版室\ 99–101页（PDF 106–108页）])[
        // 专名号
        #show emph: underline.with(offset: 0.2em, extent: -1em / 15)

        報社活版室用的字，新聞方面除了特號標題字是用木刻而外都是用鉛鑄就的……

        關於各號鉛字的大小，頭號字大約相當於_英_文的 5-Line Nonpareil（大三十點 30-Point）；二號字約等 2-Line Pica（二四點）；三號字約等於 Great Primer（十八點）；四號字約等於 Pica（十二點）；五號字約等於 Long Primer（十點）；六號字約等於 Brevier（八點）。_中__英_文相間排印的時候，這又是一個值得注意的標準。

        根據上面的點制（Point System）比較，可以求得每一_英_寸長的地位可能排的各號字數及每一平方_英_寸可能排的字數如下：

        #figure(table(
          columns: 3,
          table.header[鉛字號數][每一英寸能排字數][每一平方能切字數],
          [1], [2], [4],
          [2], [3], [9],
          [3], [4], [16],
          [4], [5], [25],
          [5], [7], [49],
          [6], [9], [81],
        ))

        這是從字的實排法（Solid）的大約計算，而且是以方體字為標準，如果每行之間加入鉛條（Leaded）（鉛條的厚薄仍然以點作標準計算），那每方寸的地位就排不了那許多字。

        ……外國字的大小是用點計算的，如像拿_美國_來說罷。他在一八八六年以來全國的鑄字公司都採用一個標準點制。所謂點的單位是等於一英寸之七十二分之一，如像六點大的字就等於十二分之一英寸高，十二點的字等於六分之一英寸，因此排字人很容易算出每一英寸能排某種鉛字多少……
      ]

      #figure(grid(
        columns: 3,
        ..range(10).map(p => image("assets/报业1936-节选.pdf", page: p + 1)),
      ))
    ],
  ),
)
