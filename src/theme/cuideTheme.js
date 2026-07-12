// src/theme/cuideTheme.js
//
// Fonte única dos tokens de tema do Cuide — os mesmos valores de
// :root / .light-mode / .mid-mode do style.css do MiContas. Qualquer
// componente que precisar de cor de tema importa daqui, em vez de
// duplicar os hex em cada arquivo.

import { useState, useEffect } from "react";

export const THEME_KEY = "cuide_theme"; // dark | mid | light

export const THEMES = {
  dark: {
    bg: "#0f0f14", card: "#16161f", border: "#2a2a3d",
    inp: "#1e1e2a", text: "#e8e8f0", text2: "#9090b0", text3: "#606080",
    accent: "#7c6af7", accentRgb: "124,106,247",
  },
  mid: {
    bg: "#363640", card: "#40404c", border: "#5c5c6c",
    inp: "#4a4a58", text: "#f5f5fa", text2: "#c7c7d6", text3: "#9797a8",
    accent: "#a394fb", accentRgb: "163,148,251",
  },
  light: {
    bg: "#f2f2f7", card: "#ffffff", border: "#d1d1d6",
    inp: "#f4f4f8", text: "#1c1c1e", text2: "#3a3a3c", text3: "#8e8e93",
    accent: "#7c6af7", accentRgb: "124,106,247",
  },
};

export const CICLO_TEMA = ["dark", "mid", "light"];
export const ROTULO_TEMA = { dark: "🌙 Escuro", mid: "🌓 Médio", light: "☀️ Claro" };

/**
 * Hook de tema do Cuide. Lê/escreve o mesmo localStorage key em todo
 * o app, então trocar o tema numa tela reflete nas outras ao navegar.
 * (Não sincroniza em tempo real entre abas abertas — se precisar
 * disso, dá pra adicionar um listener de "storage" depois.)
 */
export function useTemaCuide() {
  const [themeKey, setThemeKey] = useState(
    () => localStorage.getItem(THEME_KEY) || "dark"
  );

  useEffect(() => {
    localStorage.setItem(THEME_KEY, themeKey);
  }, [themeKey]);

  function ciclarTema() {
    setThemeKey((atual) => CICLO_TEMA[(CICLO_TEMA.indexOf(atual) + 1) % CICLO_TEMA.length]);
  }

  return { themeKey, tema: THEMES[themeKey], ciclarTema, setThemeKey };
}
