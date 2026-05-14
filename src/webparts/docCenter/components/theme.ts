import { createTheme, ITheme } from '@fluentui/react';

// Claude-inspired palette — warm off-white neutrals + Claude orange accent.
export const claudeTheme: ITheme = createTheme({
  defaultFontStyle: {
    fontFamily: '"Inter", "Segoe UI", -apple-system, BlinkMacSystemFont, Arial, sans-serif'
  },
  palette: {
    themePrimary: '#C15F3C',
    themeLighterAlt: '#FBF1ED',
    themeLighter: '#F3D8CB',
    themeLight: '#E8B49F',
    themeTertiary: '#D08368',
    themeSecondary: '#C36945',
    themeDarkAlt: '#AC5535',
    themeDark: '#92482D',
    themeDarker: '#6C3621',
    neutralLighterAlt: '#FAFAF7',
    neutralLighter: '#F4F3EE',
    neutralLight: '#ECEAE3',
    neutralQuaternaryAlt: '#DDDCD5',
    neutralQuaternary: '#D2D1CB',
    neutralTertiaryAlt: '#BFBEB7',
    neutralTertiary: '#8A8880',
    neutralSecondary: '#5C5A53',
    neutralPrimaryAlt: '#2E2C28',
    neutralPrimary: '#1A1A19',
    neutralDark: '#0F0F0E',
    black: '#000000',
    white: '#FFFFFF'
  }
});
