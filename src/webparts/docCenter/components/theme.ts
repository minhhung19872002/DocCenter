import { createTheme, ITheme } from '@fluentui/react';

// Palette xanh lá theo design "Trung tâm Tài liệu" (accent #00913f).
export const docCenterTheme: ITheme = createTheme({
  defaultFontStyle: {
    fontFamily: '"Be Vietnam Pro", "Segoe UI", system-ui, -apple-system, sans-serif'
  },
  palette: {
    themePrimary: '#00913f',
    themeLighterAlt: '#f2faf5',
    themeLighter: '#e4f4e9',
    themeLight: '#bfe4cc',
    themeTertiary: '#5cbf85',
    themeSecondary: '#17a052',
    themeDarkAlt: '#00813a',
    themeDark: '#006e31',
    themeDarker: '#005124',
    neutralLighterAlt: '#f5f7f6',
    neutralLighter: '#f5f7f2',
    neutralLight: '#e3e7de',
    neutralQuaternaryAlt: '#dde3d7',
    neutralQuaternary: '#c8d5c3',
    neutralTertiaryAlt: '#a5b0a7',
    neutralTertiary: '#94a096',
    neutralSecondary: '#68766c',
    neutralPrimaryAlt: '#2c3a30',
    neutralPrimary: '#14261b',
    neutralDark: '#0e1a13',
    black: '#000000',
    white: '#ffffff'
  }
});
