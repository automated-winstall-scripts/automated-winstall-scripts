@echo off

:: Set title
title Company Imaging

PowerShell -NoProfile -Command "& {Start-Process PowerShell -Wait -ArgumentList '-NoProfile -WindowStyle Maximized -ExecutionPolicy Bypass -File ""C:\CompanyNameImaging\DomainJoin\Join-Domain.ps1""' -Verb RunAs}";