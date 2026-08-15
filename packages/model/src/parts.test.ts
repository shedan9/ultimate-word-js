/**
 * `settings.xml` / `fontTable.xml` / `numbering.xml` 三个部件的解析。
 *
 * 前两个的期望值取自 `gongwen-01.docx` 里实际写着的东西；`numbering.xml` 这份 fixture 没有，
 * 用手写片段覆盖 —— 编号的**解引用规则**（num → abstractNum → lvl，外加 lvlOverride）
 * 是纯结构性的，不需要 Word 真值也能确定对错。真正需要真值的是编号**文字的生成**，那在 Phase 5。
 */
import { readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import { OpcPackage, parseXml, RelType } from '@uw/ooxml';
import { beforeAll, describe, expect, it } from 'vitest';
import type { FontTable } from './font-table.ts';
import { fontNameCandidates, parseFontTable } from './font-table.ts';
import { numberingLevel, parseNumbering } from './numbering.ts';
import type { DocumentSettings } from './settings.ts';
import { DEFAULT_SETTINGS, parseSettings } from './settings.ts';

const FIXTURE = new URL('../../../apps/fidelity/fixtures/gongwen-01.docx', import.meta.url);

let settings: DocumentSettings;
let fonts: FontTable;

beforeAll(() => {
  const pkg = OpcPackage.open(new Uint8Array(readFileSync(FIXTURE)));
  const part = (t: string) => {
    const n = pkg.partNameByRelType(t);
    return n === undefined ? undefined : pkg.xml(n);
  };
  settings = parseSettings(part(RelType.SETTINGS));
  fonts = parseFontTable(part(RelType.FONT_TABLE));
});

describe('settings.xml', () => {
  it('默认制表位与标点挤压 —— 两个直接影响坐标的设置', () => {
    expect(settings.defaultTabStop).toBe(420); // <w:defaultTabStop w:val="420"/>，= 21pt
    // <w:characterSpacingControl w:val="compressPunctuation"/>：行尾全角标点可压成半宽，
    // 断行位置跟着变
    expect(settings.characterSpacingControl).toBe('compressPunctuation');
  });

  it('themeFontLang 是主题东亚字体回退的语言来源', () => {
    // <w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/>
    expect(settings.themeFontLang).toEqual({ latin: 'en-US', eastAsia: 'zh-CN', bidi: '' });
  });

  it('compat 开关按裸名字收，光秃秃的元素就是 true', () => {
    expect(settings.compat.useFELayout).toBe(true); // <w:useFELayout/>
    expect(settings.compat.balanceSingleByteDoubleByteWidth).toBe(true);
    expect(settings.compat.doNotExpandShiftReturn).toBe(true);
    expect(settings.compat.spaceForUL).toBe(true);
    expect(settings.compat.没写过的开关).toBeUndefined();
  });

  it('compatSetting 里的版本号不混进布尔字典 —— 15 读成 true 会误导', () => {
    expect(settings.compatibilityMode).toBe(15);
    expect(settings.compat.compatibilityMode).toBeUndefined();
    // 值是 0/1 的才当开关收
    expect(settings.compat.overrideTableStyleFontSizeAndJustification).toBe(true);
    expect(settings.compat.useWord2013TrackBottomHyphenation).toBe(false); // w:val="0"
  });

  it('没有 settings.xml 时用规范默认值，defaultTabStop 是 720 而不是中文模板的 420', () => {
    const d = parseSettings(undefined);
    expect(d).toEqual(DEFAULT_SETTINGS);
    expect(d.defaultTabStop).toBe(720);
    expect(d.characterSpacingControl).toBe('doNotCompress');
  });

  it('结果可结构化克隆', () => {
    expect(structuredClone(settings)).toEqual(settings);
  });
});

describe('fontTable.xml', () => {
  it('altName 是本地化字体名到英文名的桥 —— 没有它，非中文系统上一款都查不到', () => {
    expect(fonts.order).toEqual(['等线', 'Times New Roman', '等线 Light', '黑体', '仿宋']);
    expect(fonts.byName.等线?.altName).toBe('DengXian');
    expect(fonts.byName.黑体?.altName).toBe('SimHei');
    // Times New Roman 本来就是英文名，没有 altName
    expect(fonts.byName['Times New Roman']?.altName).toBeUndefined();
  });

  it('查找候选按「先文档里的名字，再 altName」排 —— 反过来会在互为 altName 时挑错人', () => {
    expect(fontNameCandidates(fonts, '黑体')).toEqual(['黑体', 'SimHei']);
    expect(fontNameCandidates(fonts, 'Times New Roman')).toEqual(['Times New Roman']);
    // 字体表里没登记的字体是合法的，不是错误
    expect(fontNameCandidates(fonts, '未登记字体')).toEqual(['未登记字体']);
  });

  it('挑替代字体的依据（panose / charset / family / pitch）如实收下', () => {
    const hei = fonts.byName.黑体;
    expect(hei?.charset).toBe('86'); // 86 = GB2312 简体中文
    expect(hei?.family).toBe('modern');
    expect(hei?.pitch).toBe('fixed');
    expect(hei?.panose).toBe('02010609060101010101');
    expect(fonts.byName['Times New Roman']?.pitch).toBe('variable');
  });

  it('内嵌字体收关系 id 与混淆密钥', () => {
    const t = parseFontTable(
      parseXml(
        `<w:fonts><w:font w:name="X">
           <w:embedRegular r:id="rId1" w:fontKey="{ABC}" w:subsetted="true"/>
           <w:embedBold r:id="rId2"/>
         </w:font></w:fonts>`,
      ),
    );
    expect(t.byName.X?.embedded).toEqual([
      { style: 'regular', relId: 'rId1', fontKey: '{ABC}', subsetted: true },
      { style: 'bold', relId: 'rId2', subsetted: false },
    ]);
  });

  it('没有 fontTable.xml 时是空表，不是抛异常', () => {
    expect(parseFontTable(undefined)).toEqual({ byName: {}, order: [] });
  });
});

describe('numbering.xml', () => {
  const XML = `<w:numbering>
      <w:abstractNum w:abstractNumId="0">
        <w:multiLevelType w:val="hybridMultilevel"/>
        <w:lvl w:ilvl="0">
          <w:start w:val="1"/>
          <w:numFmt w:val="chineseCounting"/>
          <w:lvlText w:val="%1、"/>
          <w:lvlJc w:val="left"/>
          <w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr>
          <w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="1">
          <w:numFmt w:val="decimal"/>
          <w:lvlText w:val="%1.%2"/>
          <w:suff w:val="space"/>
          <w:lvlRestart w:val="0"/>
        </w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      <w:num w:numId="2">
        <w:abstractNumId w:val="0"/>
        <w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>
      </w:num>
    </w:numbering>`;

  function load() {
    const sink = createDiagnosticSink();
    return { n: parseNumbering(parseXml(XML), sink), sink };
  }

  it('级别定义收全，缺省值按规范：start=1、suffix=tab', () => {
    const { n } = load();
    const lvl0 = numberingLevel(n, 1, 0);
    expect(lvl0).toMatchObject({
      start: 1,
      numFmt: 'chineseCounting', // 不收窄成联合类型，中文编号形式有几十种
      lvlText: '%1、',
      suffix: 'tab', // 没写 w:suff 时是制表位，不是空格
      isLegal: false,
    });
    // 每级自带的段落 / 字符属性是编号那一层级联的原料（Phase 5）
    expect(lvl0?.paraProps.indent).toEqual({ left: 420, hanging: 420 });
    expect(lvl0?.runProps.fonts?.eastAsia).toBe('黑体');

    const lvl1 = numberingLevel(n, 1, 1);
    expect(lvl1).toMatchObject({ start: 1, suffix: 'space', restartAfter: 0 });
  });

  it('同一个 abstractNum 被两个 num 引用，startOverride 只改自己那份', () => {
    const { n } = load();
    expect(numberingLevel(n, 1, 0)?.start).toBe(1);
    expect(numberingLevel(n, 2, 0)?.start).toBe(5);
    // 覆盖的是起始值，其余照旧
    expect(numberingLevel(n, 2, 0)?.lvlText).toBe('%1、');
  });

  it('numId=0 是「取消编号」，不是第 0 号编号', () => {
    const { n } = load();
    expect(numberingLevel(n, 0, 0)).toBeUndefined();
  });

  it('指向不存在的 abstractNum 时报诊断，不静默给出错编号', () => {
    const sink = createDiagnosticSink();
    const n = parseNumbering(
      parseXml(`<w:numbering><w:num w:numId="9"><w:abstractNumId w:val="7"/></w:num></w:numbering>`),
      sink,
    );
    expect(numberingLevel(n, 9, 0)).toBeUndefined();
    expect(sink.list().map((d) => d.code)).toEqual(['numbering-missing-abstract']);
  });

  it('没有 numbering.xml 时是空定义', () => {
    expect(parseNumbering(undefined, createDiagnosticSink())).toEqual({ abstract: {}, instances: {} });
  });
});
