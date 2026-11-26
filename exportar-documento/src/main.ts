import { prepararDocumentoParaExportacao } from './sanitizar';
import { gerarBlobDocx, exportarParaGoogleDocs, OPCOES_DOCX_A4, URL_GOOGLE_DOCS_NOVO } from './exportar';

const api = {
  prepararDocumentoParaExportacao,
  gerarBlobDocx,
  exportarParaGoogleDocs,
  OPCOES_DOCX_A4,
  URL_GOOGLE_DOCS_NOVO,
};

(globalThis as unknown as { SeiProExportarDoc?: typeof api }).SeiProExportarDoc = api;
