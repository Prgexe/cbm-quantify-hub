import { useState } from "react";
import { Upload, Table, BarChart3, FolderOpen, Settings, Shield, GitMerge } from "lucide-react";
import { cn } from "@/lib/utils";

const tabs = [
  { id: "upload",      label: "Upload & Importação",    icon: Upload },
  { id: "consolidated",label: "Planilha Consolidada",   icon: Table },
  { id: "count-unit",  label: "Contagem por Unidade",   icon: BarChart3 },
  { id: "summary-area",label: "Resumo por Área",        icon: FolderOpen },
  { id: "merge",       label: "Mesclar Planilhas",      icon: GitMerge },
  { id: "settings",    label: "Configurações",          icon: Settings },
] as const;

export type TabId = (typeof tabs)[number]["id"];

interface AppLayoutProps {
  activeTab: TabId;
  onTabChange: (tab: TabId) => void;
  children: React.ReactNode;
}

export default function AppLayout({ activeTab, onTabChange, children }: AppLayoutProps) {
  return (
    <div className="min-h-screen flex flex-col">
      <header className="bg-primary text-primary-foreground px-4 py-3 flex items-center gap-3 shadow-md">
        <Shield className="w-7 h-7 text-accent" />
        <div>
          <h1 className="text-lg font-bold leading-tight">Gestão de Materiais — Almoxarifado</h1>
          <p className="text-xs opacity-80">CBMERJ — Corpo de Bombeiros Militar do Estado do Rio de Janeiro</p>
        </div>
      </header>

      <nav className="bg-card border-b flex overflow-x-auto">
        {tabs.map((tab) => {
          const Icon = tab.icon;
          return (
            <button
              key={tab.id}
              onClick={() => onTabChange(tab.id)}
              className={cn(
                "flex items-center gap-2 px-4 py-3 text-sm font-medium whitespace-nowrap border-b-2 transition-colors",
                activeTab === tab.id
                  ? "border-accent text-accent"
                  : "border-transparent text-muted-foreground hover:text-foreground hover:border-border"
              )}
            >
              <Icon className="w-4 h-4" />
              <span className="hidden sm:inline">{tab.label}</span>
            </button>
          );
        })}
      </nav>

      <main className="flex-1 p-4 md:p-6 max-w-[1400px] w-full mx-auto">{children}</main>
    </div>
  );
}
