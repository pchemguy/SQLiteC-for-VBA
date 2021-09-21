@echo off

icacls "%~dp0ACLLocked\*" /deny *S-1-5-32-545:M /C /L
icacls "%~dp0ACLLocked\*" /inheritance:e /deny *S-1-5-32-545:(OI)(CI)M /C /L
