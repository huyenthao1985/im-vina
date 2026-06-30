export const chartTheme = {
  dark: {
    textColor: '#E8E8F0',
    axisColor: '#B8B8C8',
    legendTextColor: '#E8E8F0',
    gridColor: 'rgba(255,255,255,0.08)',
  },
  light: {
    textColor: '#1A1A2E',
    axisColor: '#4A4A5E',
    legendTextColor: '#1A1A2E',
    gridColor: 'rgba(0,0,0,0,0.08)',
  },
  common: {
    dataLabelFontSize: 13,
    dataLabelFontWeight: 700,
    axisFontSize: 12,
    axisFontWeight: 600,
    legendFontSize: 12,
    legendFontWeight: 600,
  }
};

export function getChartLayout(
  type: 'bar' | 'donut' | 'summary',
  themeMode: 'light' | 'dark',
  custom: {
    categoryarray?: string[];
    yaxisRange?: [number, number];
    yaxis2Range?: [number, number];
  } = {}
) {
  const t = chartTheme[themeMode];
  const c = chartTheme.common;

  if (type === 'donut') {
    return {
      paper_bgcolor: 'rgba(0,0,0,0)',
      plot_bgcolor: 'rgba(0,0,0,0)',
      font: {
        family: 'Plus Jakarta Sans, sans-serif',
        color: t.textColor,
        size: c.axisFontSize
      },
      margin: { l: 20, r: 20, t: 20, b: 20 },
      legend: {
        orientation: 'v',
        x: 0.78,
        y: 0.5,
        xanchor: 'left',
        yanchor: 'middle',
        font: {
          size: c.legendFontSize,
          color: t.legendTextColor,
          weight: c.legendFontWeight === 700 ? 'bold' : 'normal'
        }
      }
    };
  }

  // Bar and Summary layouts
  const isSummary = type === 'summary';
  return {
    paper_bgcolor: 'rgba(0,0,0,0)',
    plot_bgcolor: 'rgba(0,0,0,0)',
    font: {
      family: 'Plus Jakarta Sans, sans-serif',
      color: t.textColor,
      size: c.axisFontSize,
      weight: c.axisFontWeight === 700 ? 'bold' : 'normal'
    },
    margin: { l: 50, r: 50, t: 60, b: isSummary ? 35 : 30 },
    legend: {
      orientation: 'h',
      y: 1.15,
      x: 0.5,
      xanchor: 'center',
      font: {
        size: c.legendFontSize,
        color: t.legendTextColor,
        weight: c.legendFontWeight === 700 ? 'bold' : 'normal'
      }
    },
    bargap: 0.15,
    bargroupgap: 0.02,
    xaxis: {
      type: 'category',
      categoryorder: 'array',
      categoryarray: custom.categoryarray || [],
      gridcolor: t.gridColor,
      tickfont: {
        size: c.axisFontSize,
        color: t.axisColor,
        weight: c.axisFontWeight === 700 ? 'bold' : 'normal'
      }
    },
    yaxis: {
      gridcolor: t.gridColor,
      tickfont: {
        size: c.axisFontSize,
        color: t.axisColor,
        weight: c.axisFontWeight === 700 ? 'bold' : 'normal'
      },
      range: custom.yaxisRange
    },
    yaxis2: {
      overlaying: 'y',
      side: 'right',
      gridcolor: 'rgba(0,0,0,0)',
      tickfont: {
        size: c.axisFontSize,
        color: t.axisColor,
        weight: c.axisFontWeight === 700 ? 'bold' : 'normal'
      },
      tickformat: isSummary ? undefined : '.0f',
      range: custom.yaxis2Range
    },
    hovermode: 'x unified'
  };
}
