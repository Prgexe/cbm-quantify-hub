// VARIAÇÕES
const dicionarioGraduacoes: Record<string, string> = {
  "1º SGT": "1º Sargento",
  "2º SGT": "2º Sargento",
  "3º SGT": "3º Sargento",
  "1ºSGT": "1º Sargento",
  "2ºSGT": "2º Sargento",
  "3ºSGT": "3º Sargento",
  "SUBTEN": "Subtenente",
  "SUB TEN": "Subtenente",
  "ST": "Subtenente",
  "TEN CEL": "Tenente Coronel",
  "TENCEL": "Tenente Coronel",
  "CEL": "Coronel",
  "MAJ": "Major",
  "CAP": "Capitão",
  "1º TEN": "1º Tenente",
  "2º TEN": "2º Tenente",
  "1ºTEN": "1º Tenente",
  "2ºTEN": "2º Tenente",
  "CB": "Cabo",
  "SD": "Soldado",
  // Adicione outras se precisar, como "ASP" -> "Aspirante"
};

export const padronizarGraduacao = (graduacaoBruta: string): string => {
  if (!graduacaoBruta) return "";
  
  // Limpa espaços em branco e deixa tudo maiúsculo para evitar erros de digitação (ex: " 1º Sgt ")
  const siglaLimpa = graduacaoBruta.trim().toUpperCase();
  
  // Se existir no dicionário, retorna o nome correto. Se não, retorna o que estava lá mesmo.
  return dicionarioGraduacoes[siglaLimpa] || graduacaoBruta.trim();
};