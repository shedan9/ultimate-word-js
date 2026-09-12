import type { DocumentLayout } from '@uw/layout';

/** 可让同一个稳定 run 移到另一页，验证批注跟随模型位置而不是旧页号。 */
export function annotationLayout(shift = 0, move = false): DocumentLayout {
  return {
    pages: [0, 1, 2].map((page) => {
      const run = move && page < 2 ? 1 - page : page;
      const text = ['批注跟随正文位置', '缩放与滚动保持锚定', '跨页高亮保留范围'][run] ?? '';
      const width = text.length * 300;
      return {
        index: page,
        number: page + 1,
        sectionIndex: 0,
        geometry: { width: 9000, height: 9000, content: { x: 900, y: 900, width: 7200, height: 7200 } },
        blocks: [
          {
            kind: 'paragraph',
            id: `p${run}`,
            y: shift,
            first: true,
            last: true,
            lines: [
              {
                index: 0,
                y: shift,
                line: {
                  start: 0,
                  end: text.length,
                  x: 0,
                  width,
                  height: 450,
                  baseline: 350,
                  natural: 450,
                  leaders: [],
                  isLast: true,
                  fragments: [
                    {
                      runId: `r${run}`,
                      contentIndex: 0,
                      offset: 0,
                      text,
                      x: 0,
                      width,
                      glyphX: Array.from(text, (_, i) => i * 300),
                      font: 'serif',
                      fontSize: 300,
                      script: 'eastAsia',
                      style: {
                        bold: false,
                        italic: false,
                        color: 'auto',
                        underline: 'none',
                        strike: false,
                        doubleStrike: false,
                        vertAlign: 'baseline',
                        position: 0,
                        scale: 100,
                      },
                    },
                  ],
                },
              },
            ],
          },
        ],
      };
    }),
  };
}
export const annotationStart = { nodeId: 'r0', contentIndex: 0, offset: 0 };
export const annotationEnd = { nodeId: 'r2', contentIndex: 0, offset: 8 };
