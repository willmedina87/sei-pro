import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { inflateRawSync } from 'node:zlib';
import vm from 'node:vm';
import { parseHTML } from 'linkedom';
import { prepararDocumentoParaExportacao, normalizarCor } from '../src/sanitizar';
import { gerarBlobDocx, exportarParaGoogleDocs, URL_GOOGLE_DOCS_NOVO, type HtmlToDocxFn } from '../src/exportar';

let total = 0;
async function teste(nome: string, fn: () => void | Promise<void>) {
  await fn();
  console.log(`OK ${++total}: ${nome}`);
}

/** Leitor minimo de entradas ZIP (store/deflate) para inspecionar word/document.xml do .docx real. */
function lerEntradaZip(buf: Buffer, alvo: string): string | null {
  let offset = 0;
  while (offset + 30 <= buf.length) {
    const sig = buf.readUInt32LE(offset);
    if (sig !== 0x04034b50) break;
    const metodo = buf.readUInt16LE(offset + 8);
    const tamComp = buf.readUInt32LE(offset + 18);
    const tamNome = buf.readUInt16LE(offset + 26);
    const tamExtra = buf.readUInt16LE(offset + 28);
    const nome = buf.toString('utf8', offset + 30, offset + 30 + tamNome);
    const inicioDados = offset + 30 + tamNome + tamExtra;
    const dados = buf.subarray(inicioDados, inicioDados + tamComp);
    if (nome === alvo) {
      if (metodo === 0) return dados.toString('utf8');
      if (metodo === 8) return inflateRawSync(dados).toString('utf8');
      return null;
    }
    offset = inicioDados + tamComp;
  }
  return null;
}

function carregarHtmlToDocxReal(): HtmlToDocxFn {
  const codigo = readFileSync(new URL('../../dist/js/lib/html-to-docx.browser.js', import.meta.url), 'utf8');
  const ctx: Record<string, unknown> = {
    Buffer,
    console,
    setTimeout,
    clearTimeout,
    URL,
    Blob,
    Uint8Array,
    ArrayBuffer,
    DataView,
    TextEncoder,
    TextDecoder,
    atob,
    btoa,
    XMLHttpRequest: class XMLHttpRequest { open() {} },
    location: { host: 'sei.local', href: 'https://sei.local/', protocol: 'https:' },
    crypto: globalThis.crypto,
    performance: globalThis.performance,
  };
  ctx.global = ctx;
  ctx.window = ctx;
  ctx.self = ctx;
  vm.createContext(ctx);
  vm.runInContext(codigo, ctx);
  const fn = ctx.HTMLToDOCX as HtmlToDocxFn | undefined;
  assert.equal(typeof fn, 'function', 'HTMLToDOCX deve estar disponivel no bundle browser');
  return fn!;
}

const htmlToDocxReal = carregarHtmlToDocxReal();

await teste('normalizarCor converte rgb, rgba, hex3 e nomes para #rrggbb', () => {
  assert.equal(normalizarCor('rgb(230, 230, 230)'), '#e6e6e6');
  assert.equal(normalizarCor('rgba(0, 0, 0, 0)'), null);
  assert.equal(normalizarCor('transparent'), null);
  assert.equal(normalizarCor('#abc'), '#aabbcc');
  assert.equal(normalizarCor('#0066CC'), '#0066cc');
  assert.equal(normalizarCor('yellow'), '#ffff00');
});

await teste('fixture real despacho_visualizar.html: remove <head>, <style>, <script>, .imgBgAncora e CSS sem poluir o DOCX', async () => {
  const fixture = readFileSync(new URL('../../sei-nucleo/tests/fixtures/sei41/despacho_visualizar.html', import.meta.url), 'utf8');
  const { document } = parseHTML(fixture);

  const prep = await prepararDocumentoParaExportacao(document as unknown as Document, { destino: 'docx' });

  // Nao pode haver duplo <html>/<body> nem restos de CSS/JS/UI do SEI Pro
  assert.equal((prep.htmlCompleto.match(/<html\b/gi) ?? []).length, 1);
  assert.equal((prep.htmlCompleto.match(/<body\b/gi) ?? []).length, 1);
  assert.ok(!prep.htmlCorpo.includes('<style'), 'Nao deve conter <style>');
  assert.ok(!prep.htmlCorpo.includes('<script'), 'Nao deve conter <script>');
  assert.ok(!prep.htmlCorpo.includes('CONFIGURAÇÕES DE IMPRESSÃO'), 'Nao deve conter .imgBgAncora');
  assert.ok(!prep.htmlCorpo.includes('p.Citacao'), 'Nao deve vazar CSS do SEI');
  assert.ok(!prep.htmlCorpo.includes('onclick='), 'Nao deve manter handlers onclick');
  assert.ok(prep.htmlCorpo.includes('Governo do Estado de São Paulo') || prep.htmlCorpo.includes('Governo do Estado'));
  assert.ok(prep.htmlCorpo.includes('Processo nº 99906.713-630.000032/2025-82'));

  // Conversao real para .docx e inspecao do word/document.xml
  const blob = await gerarBlobDocx(document as unknown as Document, htmlToDocxReal, { titulo: 'Despacho' });
  const zipBuf = Buffer.from(await blob.arrayBuffer());
  const xml = lerEntradaZip(zipBuf, 'word/document.xml');
  assert.ok(xml, 'word/document.xml deve existir no .docx gerado');
  assert.ok(!xml.includes('p.Citacao'), 'word/document.xml nao pode conter regras CSS');
  assert.ok(!xml.includes('imgBgAncora'), 'word/document.xml nao pode conter CSS/UI do SEI Pro');
  assert.ok(!xml.includes('document.addEventListener'), 'word/document.xml nao pode conter JS do SEI');
  assert.ok(xml.includes('99906.713-630.000032/2025-82'), 'word/document.xml deve conter o texto do despacho');
  assert.ok(xml.includes('w:w="11906"') && xml.includes('w:h="16838"'), 'Pagina deve ser A4');
});

await teste('tabelas com bordas, colspan, rowspan, larguras e sombreamento geram XML de tabela formatado no DOCX', async () => {
  const html = `<!DOCTYPE html><html><head><style>.t{width:100%}</style></head><body>
    <table border="1" style="width:100%">
      <colgroup><col width="50%"><col width="50%"></colgroup>
      <thead>
        <tr>
          <th colspan="2" bgcolor="#e6e6e6" style="text-align:center;font-weight:bold">Resumo Orcamentario</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td rowspan="2" style="width:40%;vertical-align:middle">Item 1</td>
          <td style="width:60%;text-align:right">R$ 1.000,00</td>
        </tr>
        <tr>
          <td style="width:60%;text-align:right">R$ 2.000,00</td>
        </tr>
      </tbody>
    </table>
    <table border="0" style="width:100%">
      <tr><td>Tabela de layout sem borda</td></tr>
    </table>
  </body></html>`;
  const { document } = parseHTML(html);

  const prep = await prepararDocumentoParaExportacao(document as unknown as Document, { destino: 'docx' });
  assert.ok(!prep.htmlCorpo.includes('<colgroup'), 'colgroup redundante deve ser removido');
  assert.ok(prep.htmlCorpo.includes('border-collapse:collapse'), 'tabela deve ter border-collapse');
  assert.ok(prep.htmlCorpo.includes('border-top:1px solid #000000'), 'celula de tabela com border=1 deve receber bordas explicitas');
  assert.ok(prep.htmlCorpo.includes('background-color:#e6e6e6'), 'bgcolor deve virar background-color inline');

  const blob = await gerarBlobDocx(document as unknown as Document, htmlToDocxReal);
  const zipBuf = Buffer.from(await blob.arrayBuffer());
  const xml = lerEntradaZip(zipBuf, 'word/document.xml');
  assert.ok(xml, 'word/document.xml deve existir');
  assert.ok(xml.includes('<w:tcBorders>'), 'celulas com borda devem gerar w:tcBorders');
  assert.ok(xml.includes('<w:gridSpan w:val="2"/>'), 'colspan=2 deve gerar w:gridSpan');
  assert.ok(xml.includes('<w:vMerge w:val="restart"/>'), 'rowspan=2 deve gerar w:vMerge');
  assert.ok(xml.includes('w:fill="e6e6e6"') || xml.includes('w:fill="E6E6E6"'), 'sombreamento de cabecalho deve existir');
  assert.ok(xml.includes('Resumo Orcamentario'));
});

await teste('materializa contadores CSS ::before do SEI (Item_Nivel1/2, Inciso Romano, Alinea)', async () => {
  const html = `<!DOCTYPE html><html><body>
    <p class="Item_Nivel1">DO OBJETO</p>
    <p class="Item_Nivel2">Especificacao tecnica</p>
    <p class="Item_Inciso_Romano">Primeiro inciso</p>
    <p class="Item_Inciso_Romano">Segundo inciso</p>
    <p class="Item_Alinea_Letra">Primeira alinea</p>
    <p class="Item_Alinea_Letra">Segunda alinea</p>
  </body></html>`;
  const { document } = parseHTML(html);

  const prep = await prepararDocumentoParaExportacao(document as unknown as Document, { destino: 'docx' });
  assert.ok(prep.textoPlano.includes('1. DO OBJETO'));
  assert.ok(prep.textoPlano.includes('1.1. Especificacao tecnica'));
  assert.ok(prep.textoPlano.includes('I - Primeiro inciso'));
  assert.ok(prep.textoPlano.includes('II - Segundo inciso'));
  assert.ok(prep.textoPlano.includes('a) Primeira alinea'));
  assert.ok(prep.textoPlano.includes('b) Segunda alinea'));
});

await teste('suspende modo escuro durante leitura de estilos e restaura DOM original', async () => {
  const html = `<!DOCTYPE html><html data-spro-tema="escuro" class="dark-mode"><body class="dark-mode">
    <p id="p1">Texto principal</p>
    <button data-spro-original="" type="button">Ver original</button>
  </body></html>`;
  const { document } = parseHTML(html);
  const doc = document as unknown as Document;

  // Simula getComputedStyle que responde cores escuras quando modo escuro esta ativo
  const winFake = {
    getComputedStyle(el: Element) {
      const escuro = doc.documentElement.classList.contains('dark-mode') || !doc.documentElement.hasAttribute('data-spro-original');
      return {
        display: 'block',
        textAlign: 'justify',
        fontWeight: '400',
        fontStyle: 'normal',
        textDecorationLine: 'none',
        textDecoration: 'none',
        fontSize: '16px',
        fontFamily: 'Calibri, sans-serif',
        color: escuro ? 'rgb(231, 233, 236)' : 'rgb(0, 0, 0)',
        backgroundColor: escuro ? 'rgb(26, 29, 33)' : 'rgb(255, 255, 255)',
        textIndent: '94.5px',
        marginLeft: '0px',
        marginTop: '8px',
        marginBottom: '8px',
        lineHeight: '24px',
        textTransform: 'none',
        borderTopStyle: 'none',
        borderRightStyle: 'none',
        borderBottomStyle: 'none',
        borderLeftStyle: 'none',
        borderTopWidth: '0px',
        borderRightWidth: '0px',
        borderBottomWidth: '0px',
        borderLeftWidth: '0px',
        borderTopColor: 'rgb(0, 0, 0)',
        borderRightColor: 'rgb(0, 0, 0)',
        borderBottomColor: 'rgb(0, 0, 0)',
        borderLeftColor: 'rgb(0, 0, 0)',
        verticalAlign: 'baseline',
        paddingTop: '0px',
        paddingLeft: '0px',
      };
    },
  };
  Object.defineProperty(doc, 'defaultView', { value: winFake, configurable: true });

  const prep = await prepararDocumentoParaExportacao(doc, { destino: 'docx' });

  // Cores de modo escuro NAO podem vazar para o documento exportado
  assert.ok(!prep.htmlCorpo.includes('#e7e9ec'), 'Cor de texto do modo escuro nao pode vazar');
  assert.ok(!prep.htmlCorpo.includes('#1a1d21'), 'Cor de fundo do modo escuro nao pode vazar');
  assert.ok(!prep.htmlCorpo.includes('Ver original'), 'Botao do SEI Pro deve ser removido');
  assert.ok(prep.htmlCorpo.includes('text-indent:70.9pt'), 'Recuo de primeira linha deve ser preservado em pt');

  // DOM original deve continuar no modo escuro sem alteracao
  assert.equal(doc.documentElement.classList.contains('dark-mode'), true);
  assert.equal(doc.body.classList.contains('dark-mode'), true);
  assert.equal(doc.documentElement.hasAttribute('data-spro-original'), false);
});

await teste('exportarParaGoogleDocs copia Rich Text limpo e abre docs.new', async () => {
  const html = `<!DOCTYPE html><html><head><style>body{color:red}</style></head><body>
    <p style="font-weight:bold;text-align:center">OFICIO 01/2026</p>
    <table border="1"><tr><td>Coluna A</td><td>Coluna B</td></tr></table>
  </body></html>`;
  const { document } = parseHTML(html);

  let htmlCopiado = '';
  let textoCopiado = '';
  let urlAberta = '';

  const res = await exportarParaGoogleDocs(document as unknown as Document, {
    copiar: async (h, t) => {
      htmlCopiado = h;
      textoCopiado = t;
      return true;
    },
    abrirAba: (u) => {
      urlAberta = u;
    },
  });

  assert.equal(res.copiado, true);
  assert.equal(urlAberta, URL_GOOGLE_DOCS_NOVO);
  assert.ok(htmlCopiado.includes('OFICIO 01/2026'));
  assert.ok(htmlCopiado.includes('border-collapse:collapse'));
  assert.ok(!htmlCopiado.includes('<style'));
  assert.ok(textoCopiado.includes('Coluna A'));
});

console.log(`Todos os ${total} testes de exportar-documento passaram!`);
