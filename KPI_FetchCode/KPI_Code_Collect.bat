@echo off
D:
cd \Prism_Source
for %%f in (
Core.Service.StreamRouter
Core.Library.IndexedData
Core.Service.AuditLogging
Core.Library.Sabre
Core.Library.Protocol
Core.Service.ChannelData.Listener
Core.Service.Command
Core.Service.Entitlement
Core.Service.Router
Core.Service.SurveyData
) do (
cd %%f
git pull --reb
cd ..
)