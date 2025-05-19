# TerminalMagic

TerminalMagic is a lightweight PowerShell GUI application for managing and launching SSH connections to your Linux and Windows hosts.

## Running the Script
If running powershell scripts is restricted in your environment try using this command to run the script

`powershell.exe -ExecutionPolicy Bypass -File .\terminalmagic.ps1`

## Features
- Add/edit/delete SSH hosts with categories.
- Launch SSH sessions in new Windows Terminal tabs.
- Customizable light/dark themes with JSON files.
- Persistent preferences stored locally.
- Search/filter through hosts.
- Context menu and keyboard shortcuts:
  - `Ctrl+A` → Add new host
  - `Ctrl+E` → Edit selected host
  - `Ctrl+D` → Delete selected host

## Categories

Edit the categories.csv file directly to define or update categories:
Category,Description
Uncategorized,Default fallback category
Production,Main production servers
Lab,Test VMs and experimental nodes

## Setup the Windows SSH Agent
1. `Start-Service ssh-agent`
2. `Set-Service -Name ssh-agent -StartupType Automatic`
3. `ssh-add <ssh private key>`

## File Structure
```
TerminalMagic/
├── terminalmagic.ps1        # Main script
├── hosts/hosts.csv          # Your saved host list
├── categories.csv           # host categories
├── preferences.json         # UI and theme settings
├── themes/default.json      # Customizable theme file
```
