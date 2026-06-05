import { createTheme, ITheme } from '@fluentui/react';

// DuyTan green minimal palette (brand accent #119503).
export const claudeTheme: ITheme = createTheme({
  defaultFontStyle: {
    fontFamily: '"Inter", "Segoe UI", -apple-system, BlinkMacSystemFont, Arial, sans-serif'
  },
  palette: {
    themePrimary: '#119503',
    themeLighterAlt: '#F3FBF1',
    themeLighter: '#D4F1CB',
    themeLight: '#AAE498',
    themeTertiary: '#5BC841',
    themeSecondary: '#25A813',
    themeDarkAlt: '#108503',
    themeDark: '#0B6A02',
    themeDarker: '#084F02',
    neutralLighterAlt: '#F7F9F5',
    neutralLighter: '#F0F5EC',
    neutralLight: '#E5ECDF',
    neutralQuaternaryAlt: '#D4DBCE',
    neutralQuaternary: '#C8D2C3',
    neutralTertiaryAlt: '#AEB7AA',
    neutralTertiary: '#838B7E',
    neutralSecondary: '#5F6B58',
    neutralPrimaryAlt: '#2D3329',
    neutralPrimary: '#1A1A19',
    neutralDark: '#0F0F0E',
    black: '#000000',
    white: '#FFFFFF'
  }
});
