/*
   CONFIGURAÇÃO DO SUPABASE
*/
const SUPABASE_URL = "https://enluhezbeitujlgmikxd.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVubHVoZXpiZWl0dWpsZ21pa3hkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODY0MDcwOTYsImV4cCI6MjEwMTk4MzA5Nn0.pljANF4XcutK4z0-LaIoY_9YV_fhDX2dFgv7GISlkiE";

const SENHA_ADMIN = "123456";


// --- FUNÇÕES DO SUPABASE ---

function safeJsonParse(value) {
    if (!value) return [];
    try {
        const parsed = JSON.parse(value);
        return Array.isArray(parsed) ? parsed : [];
    } catch (error) {
        return [];
    }
}

function normalizeConfirmacao(value) {
    if (value === true || value === 'true' || value === 'True' || value === 'SIM' || value === 'Sim' || value === 'sim') {
        return 'Sim';
    }
    return 'Não';
}

function getTotalPeople(conf) {
    const companionsList = safeJsonParse(conf.acompanhantes);
    const quantidade = Number(conf.quantidade) || 0;
    return 1 + Math.max(quantidade, companionsList.length);
}

async function saveConfirmation(data) {
    const nome = String(data.mainName || '').trim();
    const quantidade = Number(data.qtyGuests) || 0;
    const acompanhantes = Array.isArray(data.guestNames) ? data.guestNames : [];
    const confirmacao = 'Não';

    if (!nome) {
        throw new Error('O nome é obrigatório.');
    }

    const payloadBase = {
        nome,
        quantidade,
        acompanhantes: JSON.stringify(acompanhantes),
        created_at: new Date().toISOString()
    };

    const payloadComConfirmacao = {
        ...payloadBase,
        confirmacao
    };

    const tentarEnvio = async (payload) => {
        const response = await fetch(`${SUPABASE_URL}/rest/v1/confirmados`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_KEY,
                'Authorization': `Bearer ${SUPABASE_KEY}`,
                'Prefer': 'return=representation'
            },
            body: JSON.stringify(payload)
        });

        const bodyText = await response.text();
        let body = null;

        try {
            body = bodyText ? JSON.parse(bodyText) : null;
        } catch (error) {
            body = null;
        }

        if (!response.ok) {
            const message = body && body.message ? body.message : `Erro do Supabase (${response.status})`;
            const texto = message.toLowerCase();
            const colunaFaltando = texto.includes('confirmacao') && (
                texto.includes('does not exist') ||
                texto.includes('schema cache') ||
                texto.includes('could not find') ||
                texto.includes('column')
            );

            if (colunaFaltando) {
                throw new Error('A coluna confirmacao ainda não existe na tabela public.confirmados. No Supabase, execute:\n\nALTER TABLE public.confirmados ADD COLUMN IF NOT EXISTS confirmacao text;');
            }

            const finalMessage = texto.includes('row level security') || texto.includes('policy') || texto.includes('permission')
                ? 'A inserção foi bloqueada pelo Supabase. Verifique a política RLS da tabela confirmados ou crie uma política de INSERT permitida para o público.'
                : message;

            throw new Error(finalMessage);
        }

        return { fallback: false, message: null };
    };

    await tentarEnvio(payloadComConfirmacao);

    return true;
}

async function getConfirmations() {
    const response = await fetch(`${SUPABASE_URL}/rest/v1/confirmados?order=created_at.desc`, {
        headers: {
            'apikey': SUPABASE_KEY,
            'Authorization': `Bearer ${SUPABASE_KEY}`
        }
    });
    return await response.json();
}

async function downloadXLSX() {
    if (typeof ExcelJS === 'undefined') {
        alert('Biblioteca ExcelJS não carregada. Adicione o <script> do ExcelJS antes do script.js no HTML.');
        return;
    }

    const confirmations = await getConfirmations();
    if (!confirmations || confirmations.length === 0) {
        alert('Nada para baixar.');
        return;
    }

    const VERDE_ESCURO = 'FF2E7D47';
    const VERDE_MEDIO = 'FF66BB6A';
    const VERDE_CLARO = 'FFE8F5E9';
    const VERDE_LISTRA = 'FFF1F8F2';
    const BRANCO = 'FFFFFFFF';
    const VERDE_TEXTO = 'FF1B5E20';

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Confirmados');

    sheet.columns = [
        { width: 4 },
        { width: 28 },
        { width: 18 },
        { width: 32 },
        { width: 10 },
        { width: 14 }
    ];

    let totalPessoas = 0;

    confirmations.forEach(c => {
        totalPessoas += getTotalPeople(c);
    });

    sheet.mergeCells('A1:B1');
    sheet.getCell('A1').value = 'Confirmação de Presença';
    sheet.getCell('A1').font = { bold: true, size: 12, color: { argb: VERDE_TEXTO } };

    sheet.mergeCells('A2:D2');
    sheet.getCell('A2').value = 'Respostas da confirmação de presença';
    sheet.getCell('A2').font = { bold: true, size: 16, color: { argb: VERDE_TEXTO } };

    sheet.mergeCells('E1:F1');
    const totalCell = sheet.getCell('E1');
    totalCell.value = `Total de pessoas: ${totalPessoas}`;
    totalCell.font = { bold: true, color: { argb: BRANCO } };
    totalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERDE_ESCURO } };
    totalCell.alignment = { horizontal: 'center' };

    sheet.getCell('E2').value = `Total: ${totalPessoas}`;
    sheet.getCell('F2').value = `Pessoas: ${totalPessoas}`;
    ['E2', 'F2'].forEach(ref => {
        const cell = sheet.getCell(ref);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERDE_CLARO } };
        cell.font = { color: { argb: VERDE_TEXTO }, bold: true, size: 10 };
        cell.alignment = { horizontal: 'center' };
    });

    sheet.addRow([]);

    sheet.mergeCells('A4:F4');
    const labelCell = sheet.getCell('A4');
    labelCell.value = 'CONVIDADOS CONFIRMADOS';
    labelCell.font = { bold: true, color: { argb: BRANCO } };
    labelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERDE_ESCURO } };
    labelCell.alignment = { horizontal: 'center' };

    const headerRow = sheet.getRow(5);
    headerRow.values = ['#', 'Nome', 'Qtd. acompanhantes', 'Acompanhantes', 'Pessoas', 'Data'];
    headerRow.eachCell(cell => {
        cell.font = { bold: true, color: { argb: BRANCO } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERDE_MEDIO } };
        cell.alignment = { horizontal: 'center' };
        cell.border = { bottom: { style: 'thin', color: { argb: VERDE_ESCURO } } };
    });

    confirmations.forEach((conf, index) => {
        const date = new Date(conf.created_at).toLocaleDateString('pt-BR');
        const companionsList = safeJsonParse(conf.acompanhantes);
        const companions = companionsList.length > 0 ? companionsList.join(', ') : '-';
        const totalPessoasLinha = getTotalPeople(conf);

        const row = sheet.addRow([
            index + 1,
            conf.nome,
            conf.quantidade,
            companions,
            totalPessoasLinha,
            date
        ]);

        row.eachCell(cell => {
            cell.alignment = { horizontal: 'center' };
        });

        if (index % 2 === 1) {
            row.eachCell(cell => {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERDE_LISTRA } };
            });
        }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/octet-stream' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'confirmacoes.xlsx';
    link.click();
}

async function clearData() {
    if (!confirm('Apagar todos os dados?')) return;

    try {
        const confirmations = await getConfirmations();
        if (!confirmations || confirmations.length === 0) {
            alert('Não há dados para apagar.');
            return;
        }

        for (const conf of confirmations) {
            const response = await fetch(`${SUPABASE_URL}/rest/v1/confirmados?id=eq.${conf.id}`, {
                method: 'DELETE',
                headers: {
                    'apikey': SUPABASE_KEY,
                    'Authorization': `Bearer ${SUPABASE_KEY}`,
                    'Content-Type': 'application/json'
                }
            });

            if (!response.ok) {
                const errorText = await response.text();
                let message = 'Erro ao apagar todos os dados.';
                try {
                    const json = JSON.parse(errorText);
                    if (json && json.message) message = json.message;
                } catch (error) {
                    message = errorText || message;
                }

                if (message.toLowerCase().includes('policy') || message.toLowerCase().includes('permission') || message.toLowerCase().includes('row level security')) {
                    throw new Error('A exclusão foi bloqueada pelo Supabase. Verifique a política RLS da tabela confirmados para permitir DELETE do público.');
                }

                throw new Error(message);
            }
        }

        alert('Dados apagados!');
        window.location.reload();
    } catch (error) {
        console.error(error);
        alert(error.message || 'Erro ao apagar todos os dados.');
    }
}


// --- PÁGINA DE CONFIRMAÇÃO ---

function generateGuestFields() {
    const qtyInput = document.getElementById('qty');
    const container = document.getElementById('guestsContainer');
    if (!qtyInput || !container) return;

    container.innerHTML = '';
    for (let i = 1; i <= (parseInt(qtyInput.value) || 0); i++) {
        const input = document.createElement('input');
        input.type = 'text';
        input.placeholder = 'Nome do Acompanhante ' + i;
        input.className = 'guest-name';
        input.required = true;
        container.appendChild(input);
    }
}

// --- PÁGINA ADMIN ---

function loginAdmin() {
    const passwordInput = document.getElementById('adminPassword').value;

    if (passwordInput === SENHA_ADMIN) {
        localStorage.setItem('adminLogado', 'true');
        document.getElementById('loginScreen').style.display = 'none';
        document.getElementById('adminScreen').style.display = 'block';
        carregarLista();
    } else {
        alert('Senha incorreta!');
    }
}

async function carregarLista() {
    const listContainer = document.getElementById('listContainer');
    const confirmations = await getConfirmations();

    if (!confirmations || confirmations.length === 0) {
        listContainer.innerHTML = '<div class="empty-state">Nenhuma confirmação ainda.</div>';
        return;
    }

    let totalPessoas = 0;
    confirmations.forEach(c => {
        totalPessoas += getTotalPeople(c);
    });

    let html = `
        <div style="background:#f0f0f0; padding:15px; border-radius:5px; margin-bottom:20px; line-height:1.8;">
            <div><strong>Total de pessoas:</strong> ${totalPessoas}</div>
        </div>
        <table style="width:100%; border-collapse:collapse;">
            <tr style="background:#4CAF50; color:white;">
                <th>#</th>
                <th>Data</th>
                <th>Nome</th>
                <th>+Qtd</th>
                <th>Acompanhantes</th>
                <th>Pessoas</th>
            </tr>
    `;

    confirmations.forEach((conf, index) => {
        let date = new Date(conf.created_at).toLocaleDateString('pt-BR');
        let companionsList = safeJsonParse(conf.acompanhantes);
        let companions = companionsList.length > 0 ? companionsList.join(', ') : '-';
        const totalPessoasLinha = getTotalPeople(conf);

        html += `<tr style="border-bottom:1px solid #ddd;">
            <td>${index + 1}</td>
            <td>${date}</td>
            <td><strong>${conf.nome}</strong></td>
            <td>${conf.quantidade}</td>
            <td>${companions}</td>
            <td>${totalPessoasLinha}</td>
        </tr>`;
    });

    html += '</table>';

    listContainer.innerHTML = html;
}

// --- INICIALIZAÇÃO ---

document.addEventListener('DOMContentLoaded', function() {
    const confirmationForm = document.getElementById('confirmationForm');
    const adminLoginForm = document.getElementById('adminLoginForm');

    if (confirmationForm) {
        confirmationForm.addEventListener('submit', async function(e) {
            e.preventDefault();
            const mainName = document.getElementById('mainName').value;
            const qtyInput = document.getElementById('qty').value;
            let guestNames = [];
            document.querySelectorAll('.guest-name').forEach(input => {
                if (input.value.trim()) guestNames.push(input.value.trim());
            });

            try {
                const sucesso = await saveConfirmation({
                    mainName,
                    qtyGuests: parseInt(qtyInput) || 0,
                    guestNames
                });

                if (sucesso) {
                    alert('Confirmado com sucesso!');
                    confirmationForm.reset();
                    document.getElementById('guestsContainer').innerHTML = '';
                }
            } catch (error) {
                console.error(error);
                alert(error.message || 'Erro ao enviar.');
            }
        });
    }

    if (adminLoginForm) {
        if (localStorage.getItem('adminLogado') === 'true') {
            document.getElementById('loginScreen').style.display = 'none';
            document.getElementById('adminScreen').style.display = 'block';
            carregarLista();
        }

        adminLoginForm.addEventListener('submit', function(e) {
            e.preventDefault();
            loginAdmin();
        });
    }
});