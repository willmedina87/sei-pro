/**
 * Sanitizacao e normalizacao de documentos HTML gerados no SEI para exportacao
 * limpa em DOCX (@turbodocx/html-to-docx) e Google Docs (Rich Text).
 */

export type DestinoExportacao = 'docx' | 'gdocs';

export interface ImagemPendente {
  elemento: HTMLImageElement;
  url: string;
}

export interface ResultadoPreparacao {
  htmlCorpo: string;
  htmlCompleto: string;
  textoPlano: string;
}

export interface OpcoesPreparacao {
  destino: DestinoExportacao;
  resolverImagem?: (url: string) => Promise<string | null>;
}

const MAX_LARGURA_CONTEUDO_PX = 620;

const TAGS_REMOVER = new Set([
  'STYLE',
  'SCRIPT',
  'NOSCRIPT',
  'LINK',
  'META',
  'TITLE',
  'HEAD',
  'TEMPLATE',
  'IFRAME',
  'OBJECT',
  'EMBED',
  'AUDIO',
  'VIDEO',
  'CANVAS',
  'BUTTON',
  'INPUT',
  'SELECT',
  'TEXTAREA',
  'FORM',
]);

const SELETORES_UI_SEIPRO = [
  '.imgBgAncora',
  '[data-spro-original]',
  '[data-spro-ui]',
  '[data-spro-papel]',
  '[data-style^="seipro-"]',
  '.spro-doc-acoes',
  '.openNewTab',
  '.spro-largura',
  '.seiProToolbar',
];

const ATRIBUTOS_PERMITIDOS: Record<string, Set<string>> = {
  '*': new Set(['style']),
  A: new Set(['href', 'title']),
  IMG: new Set(['src', 'alt', 'width', 'height']),
  TD: new Set(['colspan', 'rowspan']),
  TH: new Set(['colspan', 'rowspan']),
  OL: new Set(['start', 'type']),
  UL: new Set(['type']),
  TABLE: new Set(['cellpadding', 'cellspacing']),
};

const PROPRIEDADES_DESCARTAR = new Set([
  'user-select',
  'cursor',
  'pointer-events',
  'outline',
  'box-shadow',
  'transition',
  'animation',
  'position',
  'z-index',
  'top',
  'right',
  'bottom',
  'left',
  'overflow',
  'overflow-x',
  'overflow-y',
  'word-wrap',
  'overflow-wrap',
  'white-space',
  'counter-increment',
  'counter-reset',
]);

/** Converte rgb/rgba/hex/nome simples para #rrggbb (ou null se transparente). */
export function normalizarCor(cor: string | null | undefined): string | null {
  if (!cor) return null;
  const v = cor.trim().toLowerCase();
  if (!v || v === 'transparent' || v === 'none' || v === 'inherit' || v === 'initial') return null;
  const hex3 = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(v);
  if (hex3) return `#${hex3[1]}${hex3[1]}${hex3[2]}${hex3[2]}${hex3[3]}${hex3[3]}`;
  const hex6 = /^#([0-9a-f]{6})$/.exec(v);
  if (hex6) return `#${hex6[1]}`;
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)$/.exec(v);
  if (rgb) {
    const alfa = rgb[4] !== undefined ? Number.parseFloat(rgb[4]) : 1;
    if (alfa <= 0.01) return null;
    const h = (n: string) => Math.max(0, Math.min(255, Number.parseInt(n, 10))).toString(16).padStart(2, '0');
    return `#${h(rgb[1]!)}${h(rgb[2]!)}${h(rgb[3]!)}`;
  }
  const nomes: Record<string, string> = {
    black: '#000000',
    white: '#ffffff',
    red: '#ff0000',
    blue: '#0000ff',
    green: '#008000',
    gray: '#808080',
    grey: '#808080',
    silver: '#c0c0c0',
    yellow: '#ffff00',
    navy: '#000080',
    maroon: '#800000',
  };
  return nomes[v] ?? null;
}

function pxParaPt(valorCss: string): number | null {
  const v = valorCss.trim().toLowerCase();
  if (v.endsWith('pt')) {
    const n = Number.parseFloat(v);
    return Number.isFinite(n) ? Math.round(n * 10) / 10 : null;
  }
  if (v.endsWith('px')) {
    const n = Number.parseFloat(v);
    return Number.isFinite(n) ? Math.round(n * 0.75 * 10) / 10 : null;
  }
  if (v.endsWith('mm')) {
    const n = Number.parseFloat(v);
    return Number.isFinite(n) ? Math.round(n * 2.8346 * 10) / 10 : null;
  }
  if (v.endsWith('cm')) {
    const n = Number.parseFloat(v);
    return Number.isFinite(n) ? Math.round(n * 28.346 * 10) / 10 : null;
  }
  return null;
}

function lerEstiloInline(el: Element): Map<string, string> {
  const mapa = new Map<string, string>();
  const bruto = el.getAttribute('style') ?? '';
  for (const parte of bruto.split(';')) {
    const idx = parte.indexOf(':');
    if (idx === -1) continue;
    const prop = parte.slice(0, idx).trim().toLowerCase();
    const val = parte.slice(idx + 1).trim();
    if (!prop || !val) continue;
    if (prop.startsWith('-') || PROPRIEDADES_DESCARTAR.has(prop)) continue;
    mapa.set(prop, val);
  }
  return mapa;
}

function serializarEstilo(mapa: Map<string, string>): string {
  const partes: string[] = [];
  for (const [k, v] of mapa) {
    if (v) partes.push(`${k}:${v}`);
  }
  return partes.join(';');
}

function transformarTextoDireto(el: Element, tt: string): void {
  for (let c = el.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 3 && c.nodeValue) {
      if (tt === 'uppercase') {
        c.nodeValue = c.nodeValue.toUpperCase();
      } else if (tt === 'lowercase') {
        c.nodeValue = c.nodeValue.toLowerCase();
      } else if (tt === 'capitalize') {
        c.nodeValue = c.nodeValue.replace(/(^|\s)(\S)/g, (_m, a: string, b: string) => a + b.toUpperCase());
      }
    }
  }
}

function paraRomano(num: number): string {
  const tabela: Array<[number, string]> = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'],
    [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'],
    [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  let resto = Math.max(1, num);
  let out = '';
  for (const [v, s] of tabela) {
    while (resto >= v) {
      out += s;
      resto -= v;
    }
  }
  return out;
}

function paraLetra(num: number): string {
  const idx = Math.max(1, num) - 1;
  return String.fromCharCode(97 + (idx % 26));
}

/**
 * Materializa contadores CSS (::before) padrao do SEI em texto real no clone,
 * pois Word/html-to-docx e Google Docs ignoram pseudo-elementos ::before.
 */
function materializarContadoresSei(raiz: Element): void {
  const c = {
    i1: 0, i2: 0, i3: 0, i4: 0,
    p1: 0, p2: 0, p3: 0, p4: 0,
    romano: 0,
    letra: 0,
  };

  for (const p of Array.from(raiz.querySelectorAll('p'))) {
    const cls = p.getAttribute('class') ?? '';
    if (!cls) continue;
    const classes = new Set(cls.split(/\s+/));
    let prefixo = '';

    if (classes.has('Item_Nivel1')) {
      c.i1 += 1; c.i2 = 0; c.i3 = 0; c.i4 = 0; c.romano = 0; c.letra = 0;
      prefixo = `${c.i1}. `;
    } else if (classes.has('Item_Nivel2')) {
      c.i2 += 1; c.i3 = 0; c.i4 = 0; c.romano = 0; c.letra = 0;
      prefixo = `${c.i1 || 1}.${c.i2}. `;
    } else if (classes.has('Item_Nivel3')) {
      c.i3 += 1; c.i4 = 0; c.romano = 0; c.letra = 0;
      prefixo = `${c.i1 || 1}.${c.i2 || 1}.${c.i3}. `;
    } else if (classes.has('Item_Nivel4')) {
      c.i4 += 1; c.romano = 0; c.letra = 0;
      prefixo = `${c.i1 || 1}.${c.i2 || 1}.${c.i3 || 1}.${c.i4}. `;
    } else if (classes.has('Paragrafo_Numerado_Nivel1')) {
      c.p1 += 1; c.p2 = 0; c.p3 = 0; c.p4 = 0; c.romano = 0; c.letra = 0;
      prefixo = `${c.p1}. `;
    } else if (classes.has('Paragrafo_Numerado_Nivel2')) {
      c.p2 += 1; c.p3 = 0; c.p4 = 0; c.romano = 0; c.letra = 0;
      prefixo = `${c.p1 || 1}.${c.p2}. `;
    } else if (classes.has('Paragrafo_Numerado_Nivel3')) {
      c.p3 += 1; c.p4 = 0; c.romano = 0; c.letra = 0;
      prefixo = `${c.p1 || 1}.${c.p2 || 1}.${c.p3}. `;
    } else if (classes.has('Paragrafo_Numerado_Nivel4')) {
      c.p4 += 1; c.romano = 0; c.letra = 0;
      prefixo = `${c.p1 || 1}.${c.p2 || 1}.${c.p3 || 1}.${c.p4}. `;
    } else if (classes.has('Item_Inciso_Romano')) {
      c.romano += 1; c.letra = 0;
      prefixo = `${paraRomano(c.romano)} - `;
    } else if (classes.has('Item_Alinea_Letra')) {
      c.letra += 1;
      prefixo = `${paraLetra(c.letra)}) `;
    }

    if (prefixo) {
      const textoAtual = (p.textContent ?? '').trimStart();
      if (!textoAtual.startsWith(prefixo.trim())) {
        const doc = p.ownerDocument;
        p.insertBefore(doc.createTextNode(prefixo), p.firstChild);
      }
    }
  }
}

function ehElementoUiSeiPro(el: Element): boolean {
  for (const sel of SELETORES_UI_SEIPRO) {
    try {
      if (el.matches(sel)) return true;
    } catch {
      /* seletor nao suportado */
    }
  }
  return false;
}

function suspenderModoEscuro<T>(doc: Document, fn: () => T): T {
  const html = doc.documentElement;
  const body = doc.body;
  const tinhaTema = html?.hasAttribute('data-spro-tema') ?? false;
  const tinhaOriginal = html?.hasAttribute('data-spro-original') ?? false;
  const htmlDark = html?.classList.contains('dark-mode') ?? false;
  const bodyDark = body?.classList.contains('dark-mode') ?? false;

  if (tinhaTema && !tinhaOriginal && html) html.setAttribute('data-spro-original', '');
  if (htmlDark && html) html.classList.remove('dark-mode');
  if (bodyDark && body) body.classList.remove('dark-mode');

  try {
    return fn();
  } finally {
    if (tinhaTema && !tinhaOriginal && html) html.removeAttribute('data-spro-original');
    if (htmlDark && html) html.classList.add('dark-mode');
    if (bodyDark && body) body.classList.add('dark-mode');
  }
}

function lerBordaLado(cs: CSSStyleDeclaration, lado: 'Top' | 'Right' | 'Bottom' | 'Left'): string | null {
  const estilo = cs[`border${lado}Style` as keyof CSSStyleDeclaration] as string;
  if (!estilo || estilo === 'none' || estilo === 'hidden') return null;
  const larguraStr = cs[`border${lado}Width` as keyof CSSStyleDeclaration] as string;
  const largura = Number.parseFloat(larguraStr || '0');
  if (!Number.isFinite(largura) || largura <= 0) return null;
  const cor = normalizarCor(cs[`border${lado}Color` as keyof CSSStyleDeclaration] as string) ?? '#000000';
  const px = Math.max(1, Math.round(largura * 10) / 10);
  const tipo = estilo === 'dashed' || estilo === 'dotted' || estilo === 'double' ? estilo : 'solid';
  return `${px}px ${tipo} ${cor}`;
}

function tabelaTemBordaHtml4(table: Element): boolean {
  const attr = table.getAttribute('border');
  if (!attr) return false;
  const n = Number.parseFloat(attr);
  return Number.isFinite(n) && n > 0;
}

function limparAtributos(el: Element): void {
  const tag = el.nodeName.toUpperCase();
  const permitidosTag = ATRIBUTOS_PERMITIDOS[tag];
  const permitidosGeral = ATRIBUTOS_PERMITIDOS['*']!;

  for (const attr of Array.from(el.attributes)) {
    const nome = attr.name.toLowerCase();
    if (permitidosGeral.has(nome) || (permitidosTag && permitidosTag.has(nome))) continue;
    el.removeAttribute(attr.name);
  }
}

function removerComentarios(raiz: Node): void {
  const pilha: Node[] = [raiz];
  while (pilha.length) {
    const atual = pilha.pop()!;
    for (let filho = atual.firstChild; filho; ) {
      const prox = filho.nextSibling;
      if (filho.nodeType === 8) {
        atual.removeChild(filho);
      } else if (filho.nodeType === 1) {
        pilha.push(filho);
      }
      filho = prox;
    }
  }
}

function removerLinksSemHref(raiz: Element): void {
  for (const a of Array.from(raiz.querySelectorAll('a'))) {
    const href = (a.getAttribute('href') ?? '').trim();
    if (!href || href === '#' || href.toLowerCase().startsWith('javascript:')) {
      const span = raiz.ownerDocument.createElement('span');
      const estilo = a.getAttribute('style');
      if (estilo) span.setAttribute('style', estilo);
      while (a.firstChild) span.appendChild(a.firstChild);
      a.parentNode?.replaceChild(span, a);
    }
  }
}

/**
 * Converte URLs de imagens (mesma origem / relativas) para data: URI para que
 * o DOCX e o Rich Text colado no Google Docs sejam 100% autossuficientes.
 */
async function resolverImagemPadrao(url: string): Promise<string | null> {
  if (!url || url.startsWith('data:')) return null;
  if (typeof fetch !== 'function') return null;
  try {
    const resp = await fetch(url, { credentials: 'include' });
    if (!resp.ok) return null;
    const blob = await resp.blob();
    return await new Promise<string | null>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : null);
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}

export async function prepararDocumentoParaExportacao(
  origDoc: Document,
  opcoes: OpcoesPreparacao,
): Promise<ResultadoPreparacao> {
  const origBody = origDoc.body;
  if (!origBody) {
    return {
      htmlCorpo: '',
      htmlCompleto: '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body></body></html>',
      textoPlano: '',
    };
  }

  const imagensPendentes: ImagemPendente[] = [];

  const cloneBody = suspenderModoEscuro(origDoc, () => {
    const clone = origBody.cloneNode(true) as HTMLElement;
    const win = origDoc.defaultView;
    const origEls = Array.from(origBody.querySelectorAll('*'));
    const cloneEls = Array.from(clone.querySelectorAll('*'));
    const total = Math.min(origEls.length, cloneEls.length);
    const paraRemover: Element[] = [];

    // Passo 1: pareamento 1:1 antes de qualquer remocao de nos
    for (let i = 0; i < total; i++) {
      const oe = origEls[i]!;
      const ce = cloneEls[i]!;
      const tag = oe.nodeName.toUpperCase();

      if (TAGS_REMOVER.has(tag) || ehElementoUiSeiPro(oe) || oe.hasAttribute('hidden')) {
        paraRemover.push(ce);
        continue;
      }

      const cs = win && typeof win.getComputedStyle === 'function' ? win.getComputedStyle(oe) : null;
      if (cs && cs.display === 'none') {
        paraRemover.push(ce);
        continue;
      }

      const estilo = lerEstiloInline(ce);

      // Normaliza atributos legados HTML4 (align, valign, bgcolor, width)
      const attrAlign = oe.getAttribute('align')?.toLowerCase();
      if (attrAlign && !estilo.has('text-align') && (attrAlign === 'center' || attrAlign === 'right' || attrAlign === 'justify' || attrAlign === 'left')) {
        estilo.set('text-align', attrAlign);
      }
      const attrValign = oe.getAttribute('valign')?.toLowerCase();
      if (attrValign && !estilo.has('vertical-align')) {
        estilo.set('vertical-align', attrValign);
      }
      const attrBg = normalizarCor(oe.getAttribute('bgcolor'));
      if (attrBg && !estilo.has('background-color')) {
        estilo.set('background-color', attrBg);
      }

      if (cs) {
        const ta = cs.textAlign;
        if (ta === 'center' || ta === 'right' || ta === 'justify') {
          estilo.set('text-align', ta);
        }

        const fw = cs.fontWeight;
        if (fw === 'bold' || fw === 'bolder' || Number.parseInt(fw, 10) >= 600) {
          estilo.set('font-weight', 'bold');
        }

        if (cs.fontStyle === 'italic') {
          estilo.set('font-style', 'italic');
        }

        const tdec = cs.textDecorationLine || cs.textDecoration || '';
        if (tdec.includes('underline')) estilo.set('text-decoration', 'underline');
        else if (tdec.includes('line-through')) estilo.set('text-decoration', 'line-through');

        const fsPt = pxParaPt(cs.fontSize);
        if (fsPt && fsPt >= 6 && fsPt <= 72) {
          estilo.set('font-size', `${fsPt}pt`);
        }

        const ff = cs.fontFamily;
        if (ff) {
          const primeira = ff.replace(/["']/g, '').split(',')[0]?.trim();
          if (primeira && primeira !== 'serif' && primeira !== 'sans-serif' && primeira !== 'monospace') {
            estilo.set('font-family', primeira);
          }
        }

        const col = normalizarCor(cs.color);
        if (col && col !== '#000000') {
          estilo.set('color', col);
        }

        const bg = normalizarCor(cs.backgroundColor);
        if (bg && bg !== '#ffffff') {
          estilo.set('background-color', bg);
        }

        // Propriedades de paragrafo / bloco
        if (tag === 'P' || tag === 'DIV' || tag === 'BLOCKQUOTE' || tag === 'LI' || /^H[1-6]$/.test(tag)) {
          const indentPt = pxParaPt(cs.textIndent);
          if (indentPt && indentPt > 0) {
            estilo.set('text-indent', `${indentPt}pt`);
          }
          const mlPt = pxParaPt(cs.marginLeft);
          if (mlPt && mlPt > 2) {
            estilo.set('margin-left', `${mlPt}pt`);
          }
          const mtPt = pxParaPt(cs.marginTop);
          if (mtPt !== null && mtPt >= 0 && mtPt <= 36) {
            estilo.set('margin-top', `${mtPt}pt`);
          }
          const mbPt = pxParaPt(cs.marginBottom);
          if (mbPt !== null && mbPt >= 0 && mbPt <= 36) {
            estilo.set('margin-bottom', `${mbPt}pt`);
          }
          if (cs.lineHeight && cs.lineHeight !== 'normal') {
            const lhPx = Number.parseFloat(cs.lineHeight);
            const fsPx = Number.parseFloat(cs.fontSize);
            if (Number.isFinite(lhPx) && Number.isFinite(fsPx) && fsPx > 0) {
              const ratio = Math.round((lhPx / fsPx) * 100) / 100;
              if (ratio >= 1.0 && ratio <= 2.5) {
                estilo.set('line-height', String(ratio));
              }
            }
          }
        }

        const tt = cs.textTransform;
        if (tt === 'uppercase' || tt === 'lowercase' || tt === 'capitalize') {
          transformarTextoDireto(ce, tt);
        }
      } else {
        // Fallback quando nao ha getComputedStyle (ex.: testes linkedom)
        const tt = estilo.get('text-transform');
        if (tt === 'uppercase' || tt === 'lowercase' || tt === 'capitalize') {
          transformarTextoDireto(ce, tt);
        }
      }

      // Normalizacao especifica de TABELAS
      if (tag === 'TABLE') {
        estilo.set('border-collapse', 'collapse');
        const temBordaAttr = tabelaTemBordaHtml4(oe);
        const larguraAttr = oe.getAttribute('width')?.trim();
        const larguraEstilo = estilo.get('width');

        if (!larguraEstilo) {
          if (larguraAttr) {
            estilo.set('width', /^\d+$/.test(larguraAttr) ? `${Math.min(Number(larguraAttr), MAX_LARGURA_CONTEUDO_PX)}px` : larguraAttr);
          } else if (typeof oe.getBoundingClientRect === 'function') {
            const rect = oe.getBoundingClientRect();
            if (rect.width >= 420) {
              estilo.set('width', '100%');
            } else if (rect.width > 0) {
              estilo.set('width', `${Math.round(rect.width)}px`);
            }
          }
        }

        if (cs) {
          const bTop = lerBordaLado(cs, 'Top');
          const bRight = lerBordaLado(cs, 'Right');
          const bBottom = lerBordaLado(cs, 'Bottom');
          const bLeft = lerBordaLado(cs, 'Left');
          if (bTop) estilo.set('border-top', bTop);
          if (bRight) estilo.set('border-right', bRight);
          if (bBottom) estilo.set('border-bottom', bBottom);
          if (bLeft) estilo.set('border-left', bLeft);
          if (!bTop && !bRight && !bBottom && !bLeft && !temBordaAttr) {
            estilo.set('border', 'none');
          }
        } else if (temBordaAttr && !estilo.has('border')) {
          estilo.set('border', '1px solid #000000');
        }
      } else if (tag === 'TD' || tag === 'TH') {
        const tabelaOrig = oe.closest('table');
        const temBordaTabela = tabelaOrig ? tabelaTemBordaHtml4(tabelaOrig) : false;

        // Largura da celula
        const larguraAttr = oe.getAttribute('width')?.trim();
        if (!estilo.has('width')) {
          if (larguraAttr) {
            estilo.set('width', /^\d+$/.test(larguraAttr) ? `${larguraAttr}px` : larguraAttr);
          } else if (tabelaOrig && typeof oe.getBoundingClientRect === 'function' && typeof tabelaOrig.getBoundingClientRect === 'function') {
            const cellRect = oe.getBoundingClientRect();
            const tblRect = tabelaOrig.getBoundingClientRect();
            if (cellRect.width > 0 && tblRect.width > 0) {
              const pct = Math.round((cellRect.width / tblRect.width) * 1000) / 10;
              if (pct > 0 && pct <= 100) {
                estilo.set('width', `${pct}%`);
              }
            }
          }
        }

        // Bordas da celula
        let temAlgumaBorda = false;
        if (cs) {
          const lados: Array<['Top' | 'Right' | 'Bottom' | 'Left', string]> = [
            ['Top', 'border-top'],
            ['Right', 'border-right'],
            ['Bottom', 'border-bottom'],
            ['Left', 'border-left'],
          ];
          for (const [ladoCap, prop] of lados) {
            const b = lerBordaLado(cs, ladoCap);
            if (b) {
              estilo.set(prop, b);
              temAlgumaBorda = true;
            }
          }
          const va = cs.verticalAlign;
          if (va === 'top' || va === 'middle' || va === 'bottom') {
            estilo.set('vertical-align', va);
          }
          const padTop = pxParaPt(cs.paddingTop);
          const padLeft = pxParaPt(cs.paddingLeft);
          if (padTop !== null && padTop > 0) estilo.set('padding-top', `${padTop}pt`);
          if (padLeft !== null && padLeft > 0) estilo.set('padding-left', `${padLeft}pt`);
        }

        if (!temAlgumaBorda && temBordaTabela && !estilo.has('border') && !estilo.has('border-top')) {
          estilo.set('border-top', '1px solid #000000');
          estilo.set('border-right', '1px solid #000000');
          estilo.set('border-bottom', '1px solid #000000');
          estilo.set('border-left', '1px solid #000000');
          temAlgumaBorda = true;
        }

        if (temAlgumaBorda && !estilo.has('padding') && !estilo.has('padding-left')) {
          estilo.set('padding', '4pt 6pt');
        }
      } else if (tag === 'IMG') {
        const origImg = oe as HTMLImageElement;
        const cloneImg = ce as HTMLImageElement;
        const srcBruto = origImg.getAttribute('src') ?? '';
        const srcResolvido = origImg.currentSrc || origImg.src || srcBruto;
        if (srcResolvido && !srcResolvido.startsWith('data:')) {
          imagensPendentes.push({ elemento: cloneImg, url: srcResolvido });
        }
        // Preserva dimensoes reais renderizadas, limitadas a largura da pagina A4
        let w = 0;
        let h = 0;
        if (typeof origImg.getBoundingClientRect === 'function') {
          const r = origImg.getBoundingClientRect();
          w = Math.round(r.width);
          h = Math.round(r.height);
        }
        if (!w) w = Number.parseInt(origImg.getAttribute('width') ?? '0', 10) || origImg.naturalWidth || 0;
        if (!h) h = Number.parseInt(origImg.getAttribute('height') ?? '0', 10) || origImg.naturalHeight || 0;
        if (w > MAX_LARGURA_CONTEUDO_PX && h > 0) {
          h = Math.round((h * MAX_LARGURA_CONTEUDO_PX) / w);
          w = MAX_LARGURA_CONTEUDO_PX;
        }
        if (w > 0) {
          cloneImg.setAttribute('width', String(w));
          estilo.set('width', `${w}px`);
        }
        if (h > 0) {
          cloneImg.setAttribute('height', String(h));
          estilo.set('height', `${h}px`);
        }
      }

      // Remove margens negativas (comuns em cabecalhos do SEI, quebram no Word)
      for (const mProp of ['margin-top', 'margin-bottom', 'margin-left', 'margin-right']) {
        const v = estilo.get(mProp);
        if (v && v.trim().startsWith('-')) {
          estilo.set(mProp, '0');
        }
      }

      // Normaliza cores inline ja existentes para #rrggbb
      for (const cProp of ['color', 'background-color']) {
        const v = estilo.get(cProp);
        if (v) {
          const norm = normalizarCor(v);
          if (norm) estilo.set(cProp, norm);
          else estilo.delete(cProp);
        }
      }

      const finalStyle = serializarEstilo(estilo);
      if (finalStyle) ce.setAttribute('style', finalStyle);
      else ce.removeAttribute('style');
    }

    // Passo 2: remove elementos proibidos/ocultos/UI do SEI Pro
    for (const el of paraRemover) {
      el.parentNode?.removeChild(el);
    }

    // Passo 3: materializa contadores CSS do SEI antes de limpar classes
    materializarContadoresSei(clone);

    // Passo 4: remove <colgroup> redundante que pode conflitar com tcW por celula
    for (const cg of Array.from(clone.querySelectorAll('colgroup'))) {
      cg.parentNode?.removeChild(cg);
    }

    // Passo 5: remove comentarios HTML, links falsos (onclick sem href) e atributos sujos
    removerComentarios(clone);
    removerLinksSemHref(clone);
    for (const el of Array.from(clone.querySelectorAll('*'))) {
      limparAtributos(el);
    }

    // Passo 6: limpa margens excessivas de <p> dentro de celulas de tabela
    for (const p of Array.from(clone.querySelectorAll('td p, th p'))) {
      const m = lerEstiloInline(p);
      m.set('margin-top', '0');
      m.set('margin-bottom', '2pt');
      p.setAttribute('style', serializarEstilo(m));
    }

    return clone;
  });

  // Passo 7: embute imagens externas/relativas como data: URI
  const resolver = opcoes.resolverImagem ?? resolverImagemPadrao;
  await Promise.all(
    imagensPendentes.map(async ({ elemento, url }) => {
      const dataUri = await resolver(url);
      if (dataUri) elemento.setAttribute('src', dataUri);
    }),
  );

  let htmlCorpo = cloneBody.innerHTML.trim();
  if (opcoes.destino === 'docx') {
    // O @turbodocx/html-to-docx emite &nbsp; com codificacao dupla (vira A-circunflexo no Word)
    htmlCorpo = htmlCorpo.replace(/&nbsp;/gi, ' ').replace(/\u00a0/g, ' ');
  }

  const htmlCompleto = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${htmlCorpo}</body></html>`;
  const textoPlano = (cloneBody.textContent ?? '').replace(/\u00a0/g, ' ').replace(/\n{3,}/g, '\n\n').trim();

  return { htmlCorpo, htmlCompleto, textoPlano };
}
