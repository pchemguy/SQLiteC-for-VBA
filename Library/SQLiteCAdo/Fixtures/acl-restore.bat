@echo off

icacls "%~dp0ACLLocked\*" /remove:d *S-1-5-32-545
