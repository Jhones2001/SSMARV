import { neon } from "@netlify/neon";
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, ShadingType,
} from "docx";

export const config = { path: "/api" };

const sql = neon();
const json = (d, s = 200) =>
  new Response(JSON.stringify(d), { status: s, headers: { "content-type": "application/json" } });

/* ---------- Criação automática das tabelas ---------- */
let pronto = false;
async function init() {
  if (pronto) return;
  await sql`CREATE TABLE IF NOT EXISTS filiais (id SERIAL PRIMARY KEY, nome TEXT UNIQUE NOT NULL)`;
  await sql`CREATE TABLE IF NOT EXISTS setores (id SERIAL PRIMARY KEY, nome TEXT NOT NULL, filial TEXT NOT NULL, UNIQUE (nome, filial))`;
  await sql`CREATE TABLE IF NOT EXISTS cargos (id SERIAL PRIMARY KEY, nome TEXT UNIQUE NOT NULL)`;
  await sql`CREATE TABLE IF NOT EXISTS catalogo_itens (
    id SERIAL PRIMARY KEY,
    categoria TEXT NOT NULL CHECK (categoria IN ('UNIFORME','EPI','EPI_TERMICO')),
    nome TEXT NOT NULL,
    tamanhos TEXT[] NOT NULL DEFAULT '{}',
    ativo BOOLEAN DEFAULT TRUE,
    UNIQUE (categoria, nome))`;
  await sql`CREATE TABLE IF NOT EXISTS solicitacoes (
    id SERIAL PRIMARY KEY,
    criado_em TIMESTAMPTZ DEFAULT now(),
    motivo TEXT NOT NULL,
    matricula TEXT NOT NULL,
    nome TEXT NOT NULL,
    filial TEXT NOT NULL,
    setor TEXT NOT NULL,
    cargo TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'PENDENTE'
      CHECK (status IN ('PENDENTE','EM_ANALISE','ENTREGUE','CANCELADA')),
    observacao TEXT)`;
  await sql`CREATE TABLE IF NOT EXISTS solicitacao_itens (
    id SERIAL PRIMARY KEY,
    solicitacao_id INT REFERENCES solicitacoes(id) ON DELETE CASCADE,
    categoria TEXT NOT NULL,
    item TEXT NOT NULL,
    tamanho TEXT,
    quantidade INT DEFAULT 1)`;
  await sql`CREATE INDEX IF NOT EXISTS idx_sol_status ON solicitacoes (status, filial)`;
  pronto = true;
}

const isAdmin = (req) => {
  const url = new URL(req.url);
  const p = req.headers.get("x-admin-password") || url.searchParams.get("p");
  return !!process.env.ADMIN_PASSWORD && p === process.env.ADMIN_PASSWORD;
};

/* ---------- Geração do Word (layout R-CORP-038-01) ---------- */
const cel = (texto, o = {}) =>
  new TableCell({
    width: o.w ? { size: o.w, type: WidthType.PERCENTAGE } : undefined,
    columnSpan: o.span,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    children: String(texto ?? "").split("\n").map(
      (linha) => new Paragraph({
        alignment: o.center ? AlignmentType.CENTER : AlignmentType.LEFT,
        children: [new TextRun({ text: linha, bold: !!o.bold, size: o.size || 20 })],
      })
    ),
  });
const lin = (...c) => new TableRow({ children: c });
const tab = (rows) => new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
const par = (partes, o = {}) =>
  new Paragraph({
    spacing: { before: o.before ?? 80, after: o.after ?? 80 },
    alignment: o.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    children: partes.map((p) =>
      typeof p === "string" ? new TextRun({ text: p, size: 22 }) : new TextRun({ size: 22, ...p })),
  });

async function gerarDocx(s, itens) {
  const dataBR = new Date(s.criado_em).toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });
  const uniformes = itens.filter((i) => i.categoria === "UNIFORME");
  const epis = itens.filter((i) => i.categoria === "EPI");
  const termicos = itens.filter((i) => i.categoria === "EPI_TERMICO");
  const CINZA = "D9D9D9";

  const cabecalho = tab([
    lin(
      cel("CÓD. DOC.  R-CORP-038-01", { bold: true, w: 30 }),
      cel("REGISTROS DA QUALIDADE\nTÍTULO:  SOLICITAÇÃO DE UNIFORMES E EPIs", { bold: true, center: true, w: 50 }),
      cel("REVISÃO (03)\nPágina 1/1", { w: 20 })
    ),
  ]);

  const tabelaItens = (titulo, lista) =>
    tab([
      lin(cel(titulo, { bold: true, span: 2, fill: CINZA, center: true })),
      lin(cel("ITEM", { bold: true, w: 70 }), cel("TAMANHO", { bold: true, w: 30 })),
      ...(lista.length
        ? lista.map((i) => lin(cel(i.item), cel(i.tamanho || "")))
        : [lin(cel(" "), cel(" "))]),
    ]);

  const termicosLista = ["Balaclava", "Bota térmica", "Calça térmica", "Japona Térmica", "Luva térmica", "Meia Térmica"];
  const norm = (t) => t.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const celTermico = (nome) => {
    const achado = termicos.find((t) => norm(t.item) === norm(nome) ||
      norm(t.item).includes(norm(nome.split(" ")[0])));
    return cel(`${achado ? "☒" : "☐"} ${nome}${achado && achado.tamanho ? "\nTam.: " + achado.tamanho : ""}`, { center: true, bold: true });
  };
  const outrosTermicos = termicos.filter((t) =>
    !termicosLista.some((n) => norm(t.item).includes(norm(n.split(" ")[0]))));

  const tabelaTermica = tab([
    lin(cel("DESCRIÇÃO DO ITENS - EPIs TÉRMICOS", { bold: true, span: 6, fill: CINZA, center: true })),
    lin(...termicosLista.map(celTermico)),
    ...(outrosTermicos.length
      ? [lin(cel("Outros: " + outrosTermicos.map((t) => `${t.item} (${t.tamanho || "-"})`).join(", "), { span: 6 }))]
      : []),
  ]);

  const doc = new Document({
    sections: [{
      properties: { page: { margin: { top: 720, bottom: 720, left: 720, right: 720 } } },
      children: [
        cabecalho,
        par([{ text: "SOLICITAÇÃO DE UNIFORMES E EPIs", bold: true, size: 28 }], { center: true, before: 200 }),
        par([{ text: "Data: ", bold: true }, dataBR]),
        par([
          { text: "Unidade: ", bold: true }, s.filial + "   ",
          { text: "Setor: ", bold: true }, s.setor + "   ",
          { text: "Matrícula: ", bold: true }, s.matricula + "   ",
          { text: "Solicitante: ", bold: true }, s.nome,
        ]),
        par([{ text: "Cargo: ", bold: true }, s.cargo]),
        par([{ text: "Motivo da Solicitação: ", bold: true }, s.motivo]),
        ...(s.motivo === "Extravio/Mau Uso"
          ? [par([{ text: "*Extravio/Mau Uso: Ciente do desconto em folha de pagamento pelo uso indevido do EPI/Uniforme, conforme previsto em contrato individual de trabalho assinado e legislação vigente (artigo 462 CLT).", italics: true, size: 18 }])]
          : []),
        par([{ text: "Descrição do Material", bold: true }], { before: 200 }),
        tabelaItens("DESCRIÇÃO DO ITENS - UNIFORME", uniformes),
        par([""]),
        tabelaItens("DESCRIÇÃO DO ITENS - EPI", epis),
        par([""]),
        tabelaTermica,
      ],
    }],
  });
  return Packer.toBuffer(doc);
}

/* ---------- Roteador ---------- */
export default async (req) => {
  try {
    await init();
    const url = new URL(req.url);
    const a = url.searchParams.get("a");

    // Público: dados do formulário
    if (a === "catalogo") {
      const [filiais, setores, cargos, itens] = await Promise.all([
        sql`SELECT nome FROM filiais ORDER BY nome`,
        sql`SELECT nome, filial FROM setores ORDER BY nome`,
        sql`SELECT nome FROM cargos ORDER BY nome`,
        sql`SELECT categoria, nome, tamanhos FROM catalogo_itens WHERE ativo ORDER BY nome`,
      ]);
      return json({ filiais, setores, cargos, itens });
    }

    // Público: colaborador envia solicitação
    if (a === "enviar" && req.method === "POST") {
      const b = await req.json();
      if (!b.nome || !b.matricula || !b.filial || !b.setor || !b.cargo || !b.motivo)
        return json({ erro: "Preencha todos os campos" }, 400);
      if (!Array.isArray(b.itens) || !b.itens.length)
        return json({ erro: "Selecione ao menos um item" }, 400);
      const [s] = await sql`
        INSERT INTO solicitacoes(motivo, matricula, nome, filial, setor, cargo)
        VALUES (${b.motivo}, ${String(b.matricula).trim()}, ${String(b.nome).trim()},
                ${b.filial}, ${b.setor}, ${b.cargo}) RETURNING id`;
      for (const i of b.itens) {
        await sql`INSERT INTO solicitacao_itens(solicitacao_id, categoria, item, tamanho, quantidade)
                  VALUES (${s.id}, ${i.categoria}, ${i.item}, ${i.tamanho || null}, ${i.quantidade || 1})`;
      }
      return json({ ok: true, id: s.id });
    }

    // ---- Daqui para baixo: somente admin ----
    if (!isAdmin(req)) return json({ erro: "Não autorizado" }, 401);

    if (a === "listar") {
      const status = url.searchParams.get("status") || "PENDENTE";
      const rows = await sql`
        SELECT s.*, COALESCE(json_agg(json_build_object(
          'categoria', i.categoria, 'item', i.item, 'tamanho', i.tamanho))
          FILTER (WHERE i.id IS NOT NULL), '[]') AS itens
        FROM solicitacoes s LEFT JOIN solicitacao_itens i ON i.solicitacao_id = s.id
        WHERE s.status = ${status} GROUP BY s.id ORDER BY s.criado_em DESC`;
      const resumo = await sql`
        SELECT filial, COUNT(*)::int AS total FROM solicitacoes
        WHERE status = 'PENDENTE' GROUP BY filial ORDER BY filial`;
      return json({ rows, resumo });
    }

    if (a === "status" && req.method === "POST") {
      const { id, status } = await req.json();
      await sql`UPDATE solicitacoes SET status = ${status} WHERE id = ${id}`;
      return json({ ok: true });
    }

    if (a === "importar" && req.method === "POST") {
      const { tipo, linhas } = await req.json();
      if (tipo === "organizacao") {
        for (const r of linhas) {
          const f = String(r.Filial || "").trim(), s = String(r.Setor || "").trim(), c = String(r.Cargo || "").trim();
          if (f) await sql`INSERT INTO filiais(nome) VALUES (${f}) ON CONFLICT DO NOTHING`;
          if (f && s) await sql`INSERT INTO setores(nome, filial) VALUES (${s}, ${f}) ON CONFLICT DO NOTHING`;
          if (c) await sql`INSERT INTO cargos(nome) VALUES (${c}) ON CONFLICT DO NOTHING`;
        }
      } else if (tipo === "catalogo") {
        for (const r of linhas) {
          const cat = String(r.Categoria || "").trim().toUpperCase().replace(/[\s-]+/g, "_");
          const nome = String(r.Item || "").trim();
          const tam = String(r.Tamanhos || "").split(",").map((t) => t.trim()).filter(Boolean);
          if (!["UNIFORME", "EPI", "EPI_TERMICO"].includes(cat) || !nome) continue;
          await sql`INSERT INTO catalogo_itens(categoria, nome, tamanhos)
                    VALUES (${cat}, ${nome}, ${tam})
                    ON CONFLICT (categoria, nome) DO UPDATE SET tamanhos = ${tam}, ativo = TRUE`;
        }
      } else return json({ erro: "Tipo inválido" }, 400);
      return json({ ok: true, total: linhas.length });
    }

    if (a === "docx") {
      const id = Number(url.searchParams.get("id"));
      const [s] = await sql`SELECT * FROM solicitacoes WHERE id = ${id}`;
      if (!s) return json({ erro: "Não encontrada" }, 404);
      const itens = await sql`SELECT categoria, item, tamanho FROM solicitacao_itens WHERE solicitacao_id = ${id} ORDER BY id`;
      const buf = await gerarDocx(s, itens);
      return new Response(buf, {
        headers: {
          "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          "content-disposition": `attachment; filename="R-CORP-038-01_${s.matricula}_${id}.docx"`,
        },
      });
    }

    return json({ erro: "Ação inválida" }, 400);
  } catch (e) {
    console.error(e);
    return json({ erro: "Erro interno: " + e.message }, 500);
  }
};
