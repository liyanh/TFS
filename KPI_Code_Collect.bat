@echo off
D:
cd \Prism_Source
for %%f in (
Core.Library.Protocol
Core.Library.Sabre
Core.Service.AuditLogging
Core.Service.ChannelData.Listener
Core.Service.ChannelData.Reader
Core.Service.ChannelData.Writer
Core.Service.Command
Core.Service.DemoApp
Core.Service.Entitlement
Core.Service.Listener
Core.Service.StreamRouter
Core.Service.SurveyData.Writer
Core.Service.SurveyData.Listener
Core.Service.SurveyData.Reader
) do (
echo %%f
cd %%f
git pull --reb
cd ..
)