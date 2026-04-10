import { useState } from "react";
import { DataProvider } from "@/contexts/DataContext";
import AppLayout, { TabId } from "@/components/AppLayout";
import UploadTab from "./UploadTab";
import ConsolidatedTab from "./ConsolidatedTab";
import CountByUnitTab from "./CountByUnitTab";
import SummaryByAreaTab from "./SummaryByAreaTab";
import MergeTab from "./MergeTab";
import SettingsTab from "./SettingsTab";

export default function Index() {
  const [activeTab, setActiveTab] = useState<TabId>("upload");

  return (
    <DataProvider>
      <AppLayout activeTab={activeTab} onTabChange={setActiveTab}>
        {activeTab === "upload"       && <UploadTab />}
        {activeTab === "consolidated" && <ConsolidatedTab />}
        {activeTab === "count-unit"   && <CountByUnitTab />}
        {activeTab === "summary-area" && <SummaryByAreaTab />}
        {activeTab === "merge"        && <MergeTab />}
        {activeTab === "settings"     && <SettingsTab />}
      </AppLayout>
    </DataProvider>
  );
}
