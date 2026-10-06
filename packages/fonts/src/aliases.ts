/**
 * 同一款字体的几个名字：中文 Windows 上的字体在 `name` 表里同时登记了中文名（zh-CN）与英文名（en-US），
 * Word（GDI）按**任何一个**都找得到它。文档里写哪一个取决于写它的是谁：中文版 Word 写「宋体」，
 * 英文版 Word、WPS、各种转换器常写「SimSun」。
 *
 * 为什么不靠 `fontTable.xml` 的 `w:altName`：那条桥只有中文版 Word 会搭（「宋体」的 altName 写 SimSun），
 * 反方向没人写 —— WPS 从 PDF 转出来的文件字体表里只有一个「SimSun」、没有 altName，
 * 于是随库的「宋体」度量包查不到，整份文档退到等宽近似、每一行的宽与高都是猜的，
还每份报一串 `font-missing`。
 *
 * 只收**同一个字体文件**的几个名字（查的是名字，度量一字不差）。「仿宋_GB2312」与「仿宋」不在这里：
 * 那是两个文件，度量有没有差别没量过，该走 `FontRegistry.substitute` 并报 `fallback`。
 */
export const FONT_NAME_GROUPS: readonly (readonly string[])[] = [
  ['宋体', 'SimSun'],
  ['黑体', 'SimHei'],
  ['仿宋', 'FangSong'],
  ['楷体', 'KaiTi'],
  ['等线', 'DengXian'],
  ['微软雅黑', 'Microsoft YaHei'],
  ['新宋体', 'NSimSun'],
  ['等线 Light', 'DengXian Light'],
];
