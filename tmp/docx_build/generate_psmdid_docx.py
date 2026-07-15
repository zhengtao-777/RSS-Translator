# -*- coding: utf-8 -*-
from pathlib import Path
from xml.sax.saxutils import escape
from datetime import datetime, timezone

OUT = Path(__file__).parent / "docx_package"
(OUT / "_rels").mkdir(parents=True, exist_ok=True)
(OUT / "word" / "_rels").mkdir(parents=True, exist_ok=True)
(OUT / "docProps").mkdir(parents=True, exist_ok=True)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def runs(text: str, size: int = 24, bold: bool = False, italic: bool = False) -> str:
    rpr = [
        '<w:rPr>',
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>',
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>',
    ]
    if bold:
        rpr.append('<w:b/><w:bCs/>')
    if italic:
        rpr.append('<w:i/><w:iCs/>')
    rpr.append('</w:rPr>')
    parts = []
    segs = str(text).split("\n")
    for i, seg in enumerate(segs):
        parts.append('<w:r>' + ''.join(rpr) + '<w:t xml:space="preserve">' + escape(seg) + '</w:t></w:r>')
        if i < len(segs) - 1:
            parts.append('<w:r><w:br/></w:r>')
    return ''.join(parts)


def para(text: str = "", *, size: int = 24, bold: bool = False, align: str = "both",
         first_indent: bool = True, before: int = 0, after: int = 0,
         line: int = 408, keep_next: bool = False, shading: str | None = None) -> str:
    ppr = ['<w:pPr>']
    ppr.append(f'<w:jc w:val="{align}"/>')
    if first_indent:
        ppr.append('<w:ind w:firstLine="480"/>')
    else:
        ppr.append('<w:ind w:firstLine="0"/>')
    ppr.append(f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>')
    if keep_next:
        ppr.append('<w:keepNext/>')
    if shading:
        ppr.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{shading}"/>')
    ppr.append('</w:pPr>')
    return '<w:p>' + ''.join(ppr) + runs(text, size=size, bold=bold) + '</w:p>'


def page_break() -> str:
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def tc(text: str, width: int, *, size: int = 18, bold: bool = False,
       align: str = "center", fill: str | None = None, v_align: str = "center") -> str:
    pr = [f'<w:tcW w:w="{width}" w:type="dxa"/>', f'<w:vAlign w:val="{v_align}"/>']
    if fill:
        pr.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>')
    return ('<w:tc><w:tcPr>' + ''.join(pr) + '</w:tcPr>' +
            para(text, size=size, bold=bold, align=align, first_indent=False, line=300) + '</w:tc>')


def table(rows: list[list[str]], widths: list[int], *, size: int = 18,
          header_bold: bool = True, header_fill: str | None = None) -> str:
    grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    borders = ('<w:tblBorders>'
               '<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
               '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
               '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '</w:tblBorders>')
    out = ['<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>', borders,
           '<w:tblLayout w:type="fixed"/><w:jc w:val="center"/></w:tblPr>',
           '<w:tblGrid>', grid, '</w:tblGrid>']
    for ri, row in enumerate(rows):
        out.append('<w:tr>')
        for ci, text in enumerate(row):
            align = 'left' if (ci in (1, 2, 3) and ri > 0 and len(rows[0]) == 5) else 'center'
            out.append(tc(text, widths[ci], size=size,
                          bold=(header_bold and ri == 0), align=align,
                          fill=(header_fill if ri == 0 else None)))
        out.append('</w:tr>')
    out.append('</w:tbl>')
    return ''.join(out)


def placeholder(text: str) -> str:
    borders = ('<w:tblBorders>'
               '<w:top w:val="dashed" w:sz="8" w:space="0" w:color="777777"/>'
               '<w:left w:val="dashed" w:sz="8" w:space="0" w:color="777777"/>'
               '<w:bottom w:val="dashed" w:sz="8" w:space="0" w:color="777777"/>'
               '<w:right w:val="dashed" w:sz="8" w:space="0" w:color="777777"/>'
               '</w:tblBorders>')
    return ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:jc w:val="center"/>' +
            borders + '<w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
            '<w:tr><w:trPr><w:trHeight w:val="3600" w:hRule="exact"/></w:trPr>' +
            tc(text, 9000, size=21, align='center', v_align='center') + '</w:tr></w:tbl>')


def editor_note(text: str) -> str:
    borders = ('<w:tblBorders>'
               '<w:top w:val="single" w:sz="4" w:color="999999"/>'
               '<w:left w:val="single" w:sz="4" w:color="999999"/>'
               '<w:bottom w:val="single" w:sz="4" w:color="999999"/>'
               '<w:right w:val="single" w:sz="4" w:color="999999"/>'
               '</w:tblBorders>')
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/>' + borders +
            '</w:tblPr><w:tblGrid><w:gridCol w:w="10000"/></w:tblGrid><w:tr>' +
            tc(text, 10000, size=19, align='left', fill='F2F2F2') + '</w:tr></w:tbl>')


body: list[str] = []
body.append(para('（3）倾向得分匹配与双重差分法（PSM-DID）', size=24, bold=True, align='left', first_indent=False, after=120, keep_next=True))
body.append(para('为进一步缓解核心企业供应链金融参与与供应链系统特征之间可能存在的可观测样本自选择问题，本文在工具变量法和Heckman两阶段回归之外，进一步采用倾向得分处理与双重差分相结合的方法进行检验。具体而言，本文首先依据处理发生前核心企业及其上游供应商组合的财务特征、治理特征、经营状况和供应链结构特征估计企业进入相应供应链金融处理状态的概率，并在倾向得分共同支撑区间内提高处理组与控制组的可比性；随后在相应样本中构造双重差分模型，考察核心企业供应链金融采用强度提高或首次承担供应链金融供给方角色后，供应链运营韧性、生态韧性和演进韧性是否发生显著变化。'))
body.append(para('与基准回归、工具变量法和Heckman两阶段检验的处理保持一致，本部分围绕五条已经形成显著经验事实的关系展开：一是供应链金融采用强度对供应链经营效率水平的正向影响；二是企业以供应链金融供给方参与供应链金融安排对供应链经营效率水平的负向影响；三是供应链金融采用强度对供应商关系延续程度的正向影响；四是供给方参与对供应商关系延续程度的正向影响；五是供应链金融采用强度对创新质量趋势稳定性的正向影响。对于基准回归中未形成显著结果的供给方参与与创新质量趋势稳定性关系，本部分不再开展扩展检验。'))
body.append(para('由于SCF_exp为连续变量，PSM-DID需要进一步构造明确的处理组和控制组。本文依据SCF_exp在正值样本中的分位点设定处理口径：exp75表示核心企业当年供应链金融采用强度达到正值样本第75百分位及以上，主要用于检验供应链金融采用强度对供应链经营效率水平的影响；exp50表示达到正值样本第50百分位及以上，主要用于检验其对供应商关系延续程度的影响；exp90表示达到正值样本第90百分位及以上，主要用于检验其对创新质量趋势稳定性的影响。企业首次达到相应阈值的年份被界定为处理发生年份，之后年份为处理后时期。对于SCF_focal，本文以企业首次明确作为供应链金融供给方参与供应链金融安排的年份作为处理发生年份，并分别考察其对供应链经营效率和供应商关系延续的影响。'))
body.append(para('倾向得分模型所使用的协变量原则上均取自处理发生前或滞后一期，涵盖核心企业及供应商组合的规模、年龄、资产负债率、成长性、产权性质、治理结构、市场估值和供应链组合特征等。通过避免在倾向得分方程中纳入处理后的经营、关系和创新结果，可以降低“坏控制变量”造成的识别偏误。'))

body.append(para('表6-11  PSM-DID处理口径、匹配诊断与图件排布说明', size=21, align='center', first_indent=False, before=120, after=60, keep_next=True))
rows_611 = [
    ['处理口径', '处理定义', '对应检验路径', '匹配诊断重点', '对应图件'],
    ['exp75', 'SCF_exp达到正值样本第75百分位及以上；首次达到阈值为处理年', 'SCF_exp→供应链经营效率水平', '检验较高采用强度企业与控制组在经营、财务及供应链组合特征上的共同支撑与平衡性。', '图6-1a、图6-1b、图6-5'],
    ['focal', '企业首次以供应链金融供给方参与供应链金融安排', 'SCF_focal→供应链经营效率水平；SCF_focal→供应商关系延续程度', '供给方企业通常具有更强的信用基础和生态组织能力，重点检验匹配前后企业异质性是否收敛。', '图6-2a、图6-2b、图6-6'],
    ['exp50', 'SCF_exp达到正值样本第50百分位及以上；首次达到阈值为处理年', 'SCF_exp→供应商关系延续程度', '检验中等以上采用强度企业与控制组在关系结构、供应商组合和经营基础上的可比性。', '图6-3a、图6-3b、图6-7'],
    ['exp90', 'SCF_exp达到正值样本第90百分位及以上；首次达到阈值为处理年', 'SCF_exp→创新质量趋势稳定性', '检验高采用强度企业与控制组在研发基础、创新能力和企业特征上的平衡性。', '图6-4a、图6-4b、图6-8'],
]
body.append(table(rows_611, [1100, 2500, 2400, 3000, 1400], size=17, header_fill='E7E6E6'))
body.append(para('注：各图应使用与表6-12相对应的最终估计样本和倾向得分模型生成。同一focal处理样本同时服务于供给方参与对供应链经营效率和供应商关系延续两条路径，无需重复绘制两组完全相同的匹配诊断图。', size=18, align='both', first_indent=False, line=300))

body.append(para('倾向得分分布图用于判断处理组与控制组是否具有充分的共同支撑，并考察匹配或倾向得分处理后两组在可观测特征上的可比性是否提高。正式判断应同时关注分布曲线的重叠区间、尾部观测是否缺乏共同支撑，以及匹配后曲线是否较匹配前进一步接近。图6-1a和图6-1b对应exp75口径，主要服务于供应链金融采用强度对供应链经营效率水平的PSM-DID检验。'))
body.append(editor_note('【排版说明：当前归档材料未包含可直接嵌入的第三项研究PSM诊断PNG或GPH文件，以下保留图件位置与题注。正式定稿时请用最终估计样本生成的图件替换占位框，并据实际图形修订匹配质量表述。】'))
body.append(placeholder('此处插入图6-1a\nexp75口径下匹配前倾向得分分布'))
body.append(para('图6-1a  exp75口径下匹配前倾向得分分布', size=21, align='center', first_indent=False, after=80))
body.append(placeholder('此处插入图6-1b\nexp75口径下匹配后倾向得分分布'))
body.append(para('图6-1b  exp75口径下匹配后倾向得分分布', size=21, align='center', first_indent=False, after=80))

body.append(page_break())
body.append(para('图6-2a和图6-2b对应focal口径，同时服务于供给方参与对供应链经营效率和供应商关系延续两条路径。由于供应链金融供给方企业通常具有较强的信用组织能力、平台运营能力和供应链治理能力，处理组与控制组在处理发生前可能存在较明显差异。因此，本组图应重点检验供给方企业与非供给方企业之间是否具有充分共同支撑，以及倾向得分处理能否提高两组样本的可比性。'))
body.append(placeholder('此处插入图6-2a\nfocal口径下匹配前倾向得分分布'))
body.append(para('图6-2a  focal口径下匹配前倾向得分分布', size=21, align='center', first_indent=False, after=80))
body.append(placeholder('此处插入图6-2b\nfocal口径下匹配后倾向得分分布'))
body.append(para('图6-2b  focal口径下匹配后倾向得分分布', size=21, align='center', first_indent=False, after=80))

body.append(page_break())
body.append(para('图6-3a和图6-3b对应exp50口径，主要服务于供应链金融采用强度对供应商关系延续程度的PSM-DID检验。exp50刻画供应链金融采用达到中等以上强度的处理状态，适合考察供应链金融安排能否通过账期协调、信用识别、履约保障和合作预期稳定，降低供应商关系中断的可能性。'))
body.append(placeholder('此处插入图6-3a\nexp50口径下匹配前倾向得分分布'))
body.append(para('图6-3a  exp50口径下匹配前倾向得分分布', size=21, align='center', first_indent=False, after=80))
body.append(placeholder('此处插入图6-3b\nexp50口径下匹配后倾向得分分布'))
body.append(para('图6-3b  exp50口径下匹配后倾向得分分布', size=21, align='center', first_indent=False, after=80))

body.append(page_break())
body.append(para('图6-4a和图6-4b对应exp90口径，主要服务于供应链金融采用强度对创新质量趋势稳定性的PSM-DID检验。创新质量趋势稳定性具有较强的长期性和累积性，因此采用较高强度处理口径，有助于识别供应链金融持续应用所形成的资金保障、信息协同和创新资源配置效应。'))
body.append(placeholder('此处插入图6-4a\nexp90口径下匹配前倾向得分分布'))
body.append(para('图6-4a  exp90口径下匹配前倾向得分分布', size=21, align='center', first_indent=False, after=80))
body.append(placeholder('此处插入图6-4b\nexp90口径下匹配后倾向得分分布'))
body.append(para('图6-4b  exp90口径下匹配后倾向得分分布', size=21, align='center', first_indent=False, after=80))

body.append(page_break())
body.append(para('除倾向得分分布外，本文进一步使用协变量标准化差异图检验匹配前后的平衡性。该图比较各协变量在处理组和控制组之间的绝对标准化差异；匹配后差异越接近零，说明两组在相应协变量上的可比性越高。10%的绝对标准化差异通常可作为经验性参考线，但平衡性判断仍应结合样本量、共同支撑范围和估计方法综合进行。图6-5至图6-8分别报告exp75、focal、exp50和exp90口径下的协变量标准化差异。'))
body.append(placeholder('此处插入图6-5\nexp75口径下协变量标准化差异图'))
body.append(para('图6-5  exp75口径下协变量标准化差异图', size=21, align='center', first_indent=False, after=80))
body.append(placeholder('此处插入图6-6\nfocal口径下协变量标准化差异图'))
body.append(para('图6-6  focal口径下协变量标准化差异图', size=21, align='center', first_indent=False, after=80))

body.append(page_break())
body.append(placeholder('此处插入图6-7\nexp50口径下协变量标准化差异图'))
body.append(para('图6-7  exp50口径下协变量标准化差异图', size=21, align='center', first_indent=False, after=80))
body.append(placeholder('此处插入图6-8\nexp90口径下协变量标准化差异图'))
body.append(para('图6-8  exp90口径下协变量标准化差异图', size=21, align='center', first_indent=False, after=80))

body.append(page_break())
body.append(para('在完成倾向得分处理和匹配诊断后，本文进一步在相应样本上估计双重差分模型。表6-12集中报告五条基准显著路径的PSM-DID结果。考虑到供给方参与对供应商关系延续的可用结果采用逆概率加权双重差分估计，表中同时列示具体估计方法，以避免将不同倾向得分处理方式混同。'))
body.append(para('表6-12  倾向得分处理与双重差分法（PSM-DID）检验', size=21, align='center', first_indent=False, before=120, after=60, keep_next=True))
rows_612 = [
    ['变量', '（1）\n供应链经营效率水平', '（2）\n供应链经营效率水平', '（3）\n供应商关系延续程度', '（4）\n供应商关系延续程度', '（5）\n创新质量趋势稳定性'],
    ['核心解释变量', 'SCF_exp', 'SCF_focal', 'SCF_exp', 'SCF_focal', 'SCF_exp'],
    ['DID', '0.2187*', '−0.3794*', '0.2414***', '0.2129*', '0.5477***'],
    ['标准误', '(0.1287)', '(0.1928)', '(0.0595)', '(0.1285)', '(0.1980)'],
    ['t值', '1.699', '−1.967', '4.055', '1.657', '2.766'],
    ['p值', '0.0904', '0.0501', '<0.001', '0.0988', '0.0061'],
    ['处理口径', 'exp75', 'focal', 'exp50', 'focal', 'exp90'],
    ['估计方法', 'DID', 'DID', 'DID', 'IPW-DID', 'DID'],
    ['Controls', '是', '是', '是', '是', '是'],
    ['核心企业固定效应', '是', '是', '是', '是', '是'],
    ['年份固定效应', '是', '是', '是', '是', '是'],
    ['行业×年份固定效应', '是', '是', '否', '是', '是'],
    ['地区×年份固定效应', '是', '是', '是', '否', '否'],
    ['Observations', '1,001', '1,001', '1,143', '922', '1,005'],
    ['R²', '0.8529', '0.8511', '0.5875', '0.7554', '0.6509'],
]
body.append(table(rows_612, [1700, 1740, 1740, 1740, 1740, 1740], size=16, header_fill='E7E6E6'))
body.append(para('注：***、**、*分别表示在1%、5%、10%的水平上显著，括号内为聚类至核心企业层面的稳健标准误。exp50、exp75和exp90分别表示核心企业供应链金融采用强度达到正值样本第50、第75和第90百分位及以上；focal表示核心企业首次以供应链金融供给方参与供应链金融安排。第（4）列采用基于倾向得分的逆概率加权双重差分估计，其余列采用事件型双重差分估计。第（5）列的创新质量趋势稳定性在估计前进行了1%和99%缩尾处理，但正文变量名称不保留W_等技术前缀。', size=18, align='both', first_indent=False, line=300))

body.append(para('表6-12第（1）列显示，在exp75处理口径下，DID估计系数为0.2187，并在10%水平上显著为正。这表明核心企业供应链金融采用强度首次达到较高水平后，其所对应供应链的经营效率显著提高。该结果说明，在缓解可观测样本自选择问题后，供应链金融采用强度通过资金流、订单流、结算流和交易数据的协同组织提高供应链经营连续性与资源配置效率的结论仍然成立。'))
body.append(para('第（2）列中，focal处理效应为−0.3794，并在10%水平上显著为负。该方向与基准回归一致，说明企业首次承担供应链金融供给方角色后，信用审核、风险承担、平台运营和多主体协调等新增职能可能在短期内增加组织成本和管理复杂性，使供给方角色尚未立即转化为供应链经营效率提升。因此，供给方参与对供应链运营韧性的影响具有明显的阶段性和成本约束，不能被简单等同于运营效率改善。'))
body.append(para('第（3）列显示，在exp50口径下，DID估计系数为0.2414，并在1%水平上显著为正，说明核心企业供应链金融采用达到中等以上强度后，供应商关系延续程度显著提高。供应链金融通过改善账期安排、交易预期、信用识别和履约保障，能够降低核心企业与供应商持续合作中的不确定性，从而提高供应链生态关系的稳定性。'))
body.append(para('第（4）列中，供给方参与的IPW-DID系数为0.2129，并在10%水平上显著为正。这表明在利用倾向得分权重提高供给方企业与非供给方企业可比性后，核心企业首次承担供应链金融供给方角色仍有助于提高供应商关系延续程度。该结果与供给方参与对经营效率的负向结果并不矛盾：供给方角色可能在短期内增加运营成本，但同时能够通过信用支持、交易增信和合作预期稳定增强供应链关系黏性。鉴于该结果处于10%显著性水平，其为供给方参与稳定供应商关系提供了边际显著的补充证据。'))
body.append(para('第（5）列显示，在exp90口径下，DID估计系数为0.5477，并在1%水平上显著为正，说明核心企业供应链金融采用达到较高强度后，供应链创新质量趋势稳定性显著提高。较高强度的供应链金融采用能够提高资金安排、交易协同和信息治理的连续性，为核心企业及其上游供应商持续研发和创新投入提供更稳定的经营环境，从而降低供应链创新质量发生剧烈负向波动的可能性。'))
body.append(para('总体而言，PSM-DID结果与基准回归、工具变量法和两阶段选择校正结果总体一致。供应链金融采用强度对供应链运营韧性、生态韧性和演进韧性的正向作用均得到进一步支持；供给方参与则表现出差异化影响，即其在短期内可能降低供应链经营效率，但有助于稳定供应商合作关系。由此可见，供应链金融采用强度更容易形成跨韧性维度的普遍改善，而供给方角色的作用主要体现为关系治理收益与运营组织成本之间的权衡。'))
body.append(para('需要说明的是，PSM-DID主要通过提高可观测特征上的样本可比性并控制时间不变的不可观测差异来缓解内生性问题，仍不能完全排除随时间变化且未被观测的共同冲击。因此，本文将PSM-DID结果与工具变量法、Heckman两阶段回归及后续稳健性检验结合判断，而不将其视为单一、充分的因果识别证据。'))

sect = ('<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:type="lines" w:linePitch="312"/>'
        '</w:sectPr>')

document_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>' +
                ''.join(body) + sect + '</w:body></w:document>')

styles_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:line="408" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>'''

settings_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{W_NS}"><w:zoom w:percent="100"/><w:compat/><w:defaultTabStop w:val="420"/></w:settings>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>第6章6.3.2内生性问题——PSM-DID</dc:title>
  <dc:creator>OpenAI</dc:creator>
  <cp:lastModifiedBy>OpenAI</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''

app = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><Company>OpenAI</Company><AppVersion>16.0000</AppVersion>
</Properties>'''

files = {
    OUT / '[Content_Types].xml': content_types,
    OUT / '_rels' / '.rels': rels,
    OUT / 'word' / 'document.xml': document_xml,
    OUT / 'word' / 'styles.xml': styles_xml,
    OUT / 'word' / 'settings.xml': settings_xml,
    OUT / 'word' / '_rels' / 'document.xml.rels': doc_rels,
    OUT / 'docProps' / 'core.xml': core,
    OUT / 'docProps' / 'app.xml': app,
}
for path, content in files.items():
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding='utf-8')
print(f'Generated {len(files)} package files in {OUT}')
