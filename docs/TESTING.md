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

The manifest URL carries the page key, so re-run install after deleting the helper data folder.

- [ ] The Data tab shows a SolverStudio group with Edit Data, Show Model, Solve Model, Stop, More.
- [ ] Edit Data opens the pane on Data Items. Show Model opens Model.
- [ ] The pane shows a green Gurobi 13.0.3 badge.
- [ ] Settings > Run diagnostics finishes, and "temporary sheet and names removed" passes.

## M2 solve end to end

Uses a built-in test model: capacity from A1, result into B1. Blank A1 is seeded with 10.

- [ ] Solve Model on a blank sheet opens Output and shows the Gurobi log and an Optimal chip.
- [ ] A1 is 10, B1 is 10, and the summary reads "Objective 20" and "Changed: result*".
- [ ] With A1 set to 25, Solve Model gives B1 = 25.
- [ ] With text in A1, the pane shows a clear error and B1 is unchanged.
