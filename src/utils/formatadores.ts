export const padronizarGraduacao = (graduacaoBruta: string): string => {
  if (!graduacaoBruta) return "";

  const textoLimpo = String(graduacaoBruta).trim();
  const limpa = textoLimpo.toUpperCase().replace(/[^A-Z0-9]/g, "");

  // ESPIÃO FORENSE: Vai mostrar exatamente o que sobrou da palavra
  console.log(`🕵️ Investigando: Original=[${graduacaoBruta}] -> Destruído=[${limpa}]`);

  // Defesa contra erros de digitação onde usaram "i" ou "L" no lugar do 1
  if ((limpa.includes("1") || limpa.includes("I") || limpa.includes("L")) && limpa.includes("SGT")) return "1º Sargento";
  if ((limpa.includes("2") || limpa.includes("Z")) && limpa.includes("SGT")) return "2º Sargento";
  if ((limpa.includes("3") || limpa.includes("E")) && limpa.includes("SGT")) return "3º Sargento";
  
  if (limpa.includes("SUB") || limpa === "ST") return "Subtenente";
  if (limpa.includes("TENCEL") || limpa.includes("TENCORONEL") || limpa === "TC") return "Tenente Coronel";
  if (limpa === "CEL" || limpa.includes("CORONEL")) return "Coronel";
  if (limpa === "MAJ" || limpa.includes("MAJOR")) return "Major";
  if (limpa === "CAP" || limpa.includes("CAPITAO")) return "Capitão";
  
  if ((limpa.includes("1") || limpa.includes("I") || limpa.includes("L")) && limpa.includes("TEN")) return "1º Tenente";
  if ((limpa.includes("2") || limpa.includes("Z")) && limpa.includes("TEN")) return "2º Tenente";
  
  if (limpa === "CB" || limpa.includes("CABO")) return "Cabo";
  if (limpa === "SD" || limpa.includes("SOLDADO")) return "Soldado";

  // SE CHEGAR AQUI, VAI MOSTRAR UM AVISO GRITANTE PARA PROVAR QUE O CÓDIGO RODOU
  return "ERRO_SISTEMA: " + textoLimpo;
};