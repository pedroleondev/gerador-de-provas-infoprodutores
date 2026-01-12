function getRandomChar() {
    const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    return letters[Math.floor(Math.random() * letters.length)];
}

function getRandomNumber() {
    return Math.floor(Math.random() * 90) + 10; // Número entre 10 e 99
}

// Gerando valores necessários
const number_id = getRandomNumber();
const id_transacao = `${number_id}${getRandomChar()}${getRandomChar()}${getRandomChar()}`;

// Retornando o objeto dentro de `query`
return [{
    json: {
        query: {
            cliente: $json.Cliente || "Desconhecido", // Pegando corretamente do input
            telefone: $json.chat_id,
            // number_id, // Número gerado
            id_transacao, // ID único para identificar a transação
            descricao: $json.descricao || "",
            tipo: $json.tipo || "DESPESA",
            categoria: $json.categoria || "Outros",
            conta: $json.conta || "Conta pessoal",
            valor: $json.valor,
            pago: $json.pago || false, // Se não existir, assume false
            data_transacao: $json.data || new Date().toISOString().split("T")[0] // Data de hoje se não existir
        }
    }
}];
