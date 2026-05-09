import { readFileSync, writeFileSync, existsSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { createHash } from "crypto";

const __dirname = dirname(fileURLToPath(import.meta.url));
const DB_PATH = join(__dirname, "data.json");

export function hashPw(pw) {
  return createHash("sha256").update(pw).digest("hex");
}

function defaultData() {
  const pw = hashPw("demo123");
  return {
    nextId: 100,
    users: [
      {id:1,name:"Oscar Avila",email:"oscar.avila@hannover-re.com",cc:"79749093",phone:"+57 310 555 1234",role:"employee",area:"Administracion",password:pw},
      {id:2,name:"Juan Medina",email:"juan.medina@hannover-re.com",cc:"80123456",phone:"+57 311 555 5678",role:"analyst",area:"Underwriting",password:pw},
      {id:3,name:"Carolina Lopez",email:"carolina.lopez@hannover-re.com",cc:"52987654",phone:"+57 312 555 9012",role:"team_leader",area:"Underwriting",password:pw},
      {id:4,name:"Maria Uribe",email:"maria.uribe@hannover-re.com",cc:"51456789",phone:"+57 313 555 3456",role:"admin",area:"Administracion",password:pw},
      {id:5,name:"Andrea Ruiz",email:"andrea.ruiz@hannover-re.com",cc:"53112233",phone:"+57 314 555 7890",role:"accountant",area:"Finanzas",password:pw},
      {id:6,name:"Miguel Guarin",email:"miguel.guarin@hannover-re.com",cc:"79876543",phone:"+57 315 555 2345",role:"general_manager",area:"Gerencia",password:pw},
      {id:7,name:"Laura Martinez",email:"laura.martinez@hannover-re.com",cc:"52234567",phone:"+57 316 555 6789",role:"compliance",area:"Compliance",password:pw},
      {id:8,name:"Diego Torres",email:"diego.torres@hannover-re.com",cc:"80345678",phone:"+57 317 555 0123",role:"auditor",area:"Auditoria",password:pw},
    ],
    reports: [
      {id:1,code:"RPT-2025-001",employee_id:1,employee_name:"Oscar Avila",area:"Administracion",cc:"79749093",destination:"Mexico",trip_purpose:"Visita clientes",date_from:"2025-05-25",date_to:"2025-06-07",presentation_date:"2025-06-18",currency:"USD",cost_center:"401",type:"anticipo",status:"approved_leader",current_step:3,created_at:"2025-06-18",
        expenses:[
          {id:1,date:"2025-05-25",currency:"USD",category:"alojamiento",amount:890,obs:"Hotel Ciudad de Mexico - 7 noches",has_receipt:1,attachments:[]},
          {id:2,date:"2025-05-24",currency:"COP",category:"alimentacion",amount:20000,obs:"Cafe aeropuerto El Dorado",has_receipt:1,attachments:[]},
          {id:3,date:"2025-05-25",currency:"USD",category:"transporte",amount:200,obs:"Uber Bogota y Mexico",has_receipt:1,attachments:[]},
          {id:4,date:"2025-05-26",currency:"USD",category:"invitaciones",amount:85,obs:"Cena cliente Seguros Atlas",has_receipt:1,attachments:[]},
          {id:5,date:"2025-05-28",currency:"MXN",category:"alimentacion",amount:450,obs:"Almuerzo zona Rosa",has_receipt:1,attachments:[]},
          {id:6,date:"2025-05-30",currency:"USD",category:"otros",amount:35,obs:"Lavanderia hotel",has_receipt:0,attachments:[]},
        ],
        approvals:[{step:1,by_name:"Oscar Avila",date:"2025-06-18",comment:""},{step:2,by_name:"Carolina Lopez",date:"2025-06-19",comment:"Aprobado."}]
      },
      {id:2,code:"RPT-2025-002",employee_id:2,employee_name:"Juan Medina",area:"Underwriting",cc:"80123456",destination:"Medellin",trip_purpose:"Evento SURA",date_from:"2025-11-18",date_to:"2025-11-20",presentation_date:"2025-12-18",currency:"COP",cost_center:"301",type:"tarjeta",status:"submitted",current_step:2,created_at:"2025-12-18",
        expenses:[
          {id:7,date:"2025-11-18",currency:"COP",category:"transporte",amount:45200,obs:"Uber Aeropuerto MDE",has_receipt:1,attachments:[]},
          {id:8,date:"2025-11-18",currency:"COP",category:"alojamiento",amount:380000,obs:"Hotel El Tesoro - 2 noches",has_receipt:1,attachments:[]},
          {id:9,date:"2025-11-19",currency:"COP",category:"invitaciones",amount:250000,obs:"Invitacion SURA evento",has_receipt:1,attachments:[]},
          {id:10,date:"2025-11-20",currency:"COP",category:"transporte",amount:52300,obs:"Uber a Aeropuerto",has_receipt:1,attachments:[]},
          {id:11,date:"2025-11-20",currency:"COP",category:"tiquete",amount:485000,obs:"Avianca BOG-MDE-BOG",has_receipt:1,attachments:[]},
        ],
        approvals:[{step:1,by_name:"Juan Medina",date:"2025-12-18",comment:""}]
      },
      {id:3,code:"RPT-2025-003",employee_id:2,employee_name:"Juan Medina",area:"Underwriting",cc:"80123456",destination:"Cartagena",trip_purpose:"Congreso Fasecolda",date_from:"2025-09-10",date_to:"2025-09-12",presentation_date:"2025-09-17",currency:"COP",cost_center:"301",type:"anticipo",status:"approved_final",current_step:5,created_at:"2025-09-17",
        expenses:[
          {id:12,date:"2025-09-10",currency:"COP",category:"tiquete",amount:520000,obs:"LATAM BOG-CTG-BOG",has_receipt:1,attachments:[]},
          {id:13,date:"2025-09-10",currency:"COP",category:"transporte",amount:38000,obs:"Uber a hotel",has_receipt:1,attachments:[]},
          {id:14,date:"2025-09-10",currency:"COP",category:"alojamiento",amount:450000,obs:"Hilton Cartagena 2n",has_receipt:1,attachments:[]},
          {id:15,date:"2025-09-11",currency:"COP",category:"alimentacion",amount:95000,obs:"Almuerzo y cena",has_receipt:1,attachments:[]},
          {id:16,date:"2025-09-12",currency:"COP",category:"transporte",amount:42000,obs:"Uber a aeropuerto",has_receipt:1,attachments:[]},
        ],
        approvals:[{step:1,by_name:"Juan Medina",date:"2025-09-17",comment:""},{step:2,by_name:"Carolina Lopez",date:"2025-09-18",comment:"OK"},{step:3,by_name:"Maria Uribe",date:"2025-09-19",comment:"Verificado"},{step:4,by_name:"Andrea Ruiz",date:"2025-09-20",comment:"Contabilizado"},{step:5,by_name:"Miguel Guarin",date:"2025-09-21",comment:"Aprobado"}]
      },
    ],
    currencies: [
      {code:"USD",name:"Dolar EE.UU.",rate_to_cop:4250,symbol:"$",is_custom:0},
      {code:"EUR",name:"Euro",rate_to_cop:4600,symbol:"\u20AC",is_custom:0},
      {code:"COP",name:"Peso Colombiano",rate_to_cop:1,symbol:"$",is_custom:0},
      {code:"MXN",name:"Peso Mexicano",rate_to_cop:248,symbol:"MXN",is_custom:0},
      {code:"BRL",name:"Real Brasileno",rate_to_cop:860,symbol:"R$",is_custom:0},
      {code:"ARS",name:"Peso Argentino",rate_to_cop:3.8,symbol:"ARS",is_custom:0},
      {code:"CLP",name:"Peso Chileno",rate_to_cop:4.5,symbol:"CLP",is_custom:0},
      {code:"PEN",name:"Sol Peruano",rate_to_cop:1142,symbol:"S/",is_custom:0},
      {code:"GBP",name:"Libra Esterlina",rate_to_cop:5380,symbol:"\u00A3",is_custom:0},
      {code:"CAD",name:"Dolar Canadiense",rate_to_cop:3132,symbol:"C$",is_custom:0},
      {code:"CHF",name:"Franco Suizo",rate_to_cop:4830,symbol:"CHF",is_custom:0},
    ],
  };
}

class JsonDB {
  constructor() {
    if (existsSync(DB_PATH)) {
      this.data = JSON.parse(readFileSync(DB_PATH, "utf-8"));
    } else {
      this.data = defaultData();
      this.save();
    }
  }
  save() { writeFileSync(DB_PATH, JSON.stringify(this.data, null, 2)); }
  genId() { this.data.nextId++; this.save(); return this.data.nextId; }

  // Users
  getUsers() { return this.data.users.map(({ password, ...u }) => u); }
  getUserByEmail(email) { return this.data.users.find(u => u.email?.toLowerCase() === email.toLowerCase()); }
  getUserById(id) { return this.data.users.find(u => u.id === id); }
  createUser(u) { const id = this.genId(); const nu = { id, ...u, password: hashPw("demo123") }; this.data.users.push(nu); this.save(); const { password, ...safe } = nu; return safe; }
  updateUser(id, updates) { const i = this.data.users.findIndex(u => u.id === id); if (i < 0) return null; Object.assign(this.data.users[i], updates); this.save(); const { password, ...safe } = this.data.users[i]; return safe; }
  deleteUser(id) { this.data.users = this.data.users.filter(u => u.id !== id); this.save(); }

  // Reports
  getReports() { return this.data.reports; }
  getReport(id) { return this.data.reports.find(r => r.id === id); }
  createReport(r) { const id = this.genId(); const code = "RPT-" + new Date().getFullYear() + "-" + String(Math.floor(Math.random() * 900) + 100); const nr = { id, code, ...r, expenses: (r.expenses || []).map((e, i) => ({ ...e, id: this.genId(), attachments: [] })), approvals: r.approvals || [], created_at: new Date().toISOString().split("T")[0] }; this.data.reports.push(nr); this.save(); return nr; }
  updateReport(id, updates) { const i = this.data.reports.findIndex(r => r.id === id); if (i < 0) return null; const r = this.data.reports[i]; if (updates.expenses) { updates.expenses = updates.expenses.map(e => ({ ...e, id: e.id || this.genId(), attachments: e.attachments || [] })); } Object.assign(r, updates); this.save(); return r; }
  addApproval(reportId, approval) { const r = this.getReport(reportId); if (!r) return null; r.approvals.push(approval); this.save(); return r; }

  // Attachments
  addAttachment(expenseId, att) { for (const r of this.data.reports) { const e = r.expenses.find(x => x.id === expenseId); if (e) { const id = this.genId(); e.attachments.push({ id, ...att }); e.has_receipt = 1; this.save(); return e.attachments; } } return null; }
  getAttachment(id) { for (const r of this.data.reports) { for (const e of r.expenses) { const a = (e.attachments || []).find(x => x.id === id); if (a) return a; } } return null; }

  // Currencies
  getCurrencies() { return this.data.currencies; }
  addCurrency(c) { if (this.data.currencies.find(x => x.code === c.code)) throw new Error("Ya existe"); this.data.currencies.push({ ...c, is_custom: 1 }); this.save(); }
  removeCurrency(code) { this.data.currencies = this.data.currencies.filter(c => !(c.code === code && c.is_custom)); this.save(); }
  updateTRM(rates) { this.data.currencies.forEach(c => { if (rates[c.code] !== undefined) c.rate_to_cop = rates[c.code]; }); this.save(); }
}

export const db = new JsonDB();
