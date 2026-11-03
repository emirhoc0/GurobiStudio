# Manual checks in Excel for Mac

Automated tests can't click inside Excel. These stay **unverified** until someone runs them.

## Reload after a manifest change

1. Quit Excel.
2. Run `scripts/install.sh`.
3. Delete the contents of `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`, if it exists.
   Otherwise delete `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/`.
4. Re-run `scripts/install.sh`, since clearing can also remove the sideloaded manifest.
5. Open Excel, then Home > Add-ins > GurobiStudio.

To inspect the task pane: `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`,
then right-click in the pane and choose Inspect Element.

## M1 pipeline

- [ ] The Data tab shows a SolverStudio group with Edit Data, Show Model, Solve Model, Stop, More.
- [ ] Edit Data opens the pane on Data Items. Show Model opens Model.
- [ ] The pane shows a green Gurobi 13.0.3 badge.
- [ ] Settings > Run diagnostics finishes, and the blank workbook has no leftover sheets or names.
