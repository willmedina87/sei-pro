/**
 * Geracao de Blob DOCX (A4, margens SEI, tabelas limpas) e exportacao
 * formatada para Google Docs (Rich Text na area de transferencia + docs.new).
 */

import { prepararDocumentoParaExportacao, type OpcoesPreparacao } from './sanitizar';

export type HtmlToDocxFn = (
  html: string,
  headerHtml: string | null,
  options: Record<string, unknown>,
) => Promise<Blob | ArrayBuffer | Uint8Array>;

export interface OpcoesDocx {
  titulo?: string;
  resolverImagem?: OpcoesPreparacao['resolverImagem'];
}

export interface OpcoesGoogleDocs {
  abrirAba?: (url: string) => unknown;
  copiar?: (html: string, textoPlano: string) => Promise<boolean>;
  resolverImagem?: OpcoesPreparacao['resolverImagem'];
}

export interface ResultadoGoogleDocs {
  copiado: boolean;
  htmlCompleto: string;
  textoPlano: string;
}

export const URL_GOOGLE_DOCS_NOVO = 'https://docs.new';

export const OPCOES_DOCX_A4: Record<string, unknown> = {
  orientation: 'portrait',
  font: 'Calibri',
  fontSize: 24,
  pageSize: {
    width: 11906,
    height: 16838,
  },
  margins: {
    top: 1134,
    right: 1134,
    bottom: 1134,
    left: 1417,
  },
  table: {
    row: { cantSplit: true },
    addSpacingAfter: true,
  },
  decodeUnicode: true,
};

export async function gerarBlobDocx(
  origDoc: Document,
  htmlToDocx: HtmlToDocxFn,
  opcoes: OpcoesDocx = {},
): Promise<Blob> {
  const { htmlCompleto } = await prepararDocumentoParaExportacao(origDoc, {
    destino: 'docx',
    resolverImagem: opcoes.resolverImagem,
  });

  const config: Record<string, unknown> = {
    ...OPCOES_DOCX_A4,
    ...(opcoes.titulo ? { title: opcoes.titulo } : {}),
  };

  const resultado = await htmlToDocx(htmlCompleto, null, config);
  if (resultado instanceof Blob) return resultado;
  return new Blob([resultado as BlobPart], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

async function copiarRichTextNavegador(html: string, textoPlano: string): Promise<boolean> {
  if (
    typeof navigator !== 'undefined' &&
    navigator.clipboard &&
    typeof navigator.clipboard.write === 'function' &&
    typeof ClipboardItem !== 'undefined'
  ) {
    try {
      const item = new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob([textoPlano], { type: 'text/plain' }),
      });
      await navigator.clipboard.write([item]);
      return true;
    } catch {
      /* fallback para execCommand abaixo */
    }
  }

  if (typeof document === 'undefined' || !document.body) return false;
  const container = document.createElement('div');
  container.setAttribute('contenteditable', 'true');
  container.style.cssText = 'position:fixed;left:-99999px;top:0;opacity:0;pointer-events:none;';
  container.innerHTML = html;
  document.body.appendChild(container);

  try {
    const sel = window.getSelection();
    if (!sel) return false;
    const range = document.createRange();
    range.selectNodeContents(container);
    sel.removeAllRanges();
    sel.addRange(range);
    const ok = document.execCommand('copy');
    sel.removeAllRanges();
    return ok;
  } catch {
    return false;
  } finally {
    container.remove();
  }
}

export async function exportarParaGoogleDocs(
  origDoc: Document,
  opcoes: OpcoesGoogleDocs = {},
): Promise<ResultadoGoogleDocs> {
  const { htmlCompleto, textoPlano } = await prepararDocumentoParaExportacao(origDoc, {
    destino: 'gdocs',
    resolverImagem: opcoes.resolverImagem,
  });

  const copiar = opcoes.copiar ?? copiarRichTextNavegador;
  const copiado = await copiar(htmlCompleto, textoPlano);

  const abrirAba =
    opcoes.abrirAba ??
    ((url: string) => {
      if (typeof window !== 'undefined' && typeof window.open === 'function') {
        return window.open(url, '_blank', 'noopener');
      }
      return null;
    });

  if (copiado) {
    abrirAba(URL_GOOGLE_DOCS_NOVO);
  }

  return { copiado, htmlCompleto, textoPlano };
}
