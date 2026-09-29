import type { Theme } from 'vitepress';
import DefaultTheme from 'vitepress/theme';
import DocxDemo from '../components/DocxDemo.vue';
import MermaidDiagram from '../components/MermaidDiagram.vue';
import './custom.css';

export default {
  extends: DefaultTheme,
  enhanceApp({ app }) {
    app.component('DocxDemo', DocxDemo);
    app.component('MermaidDiagram', MermaidDiagram);
  },
} satisfies Theme;
