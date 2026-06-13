# Automatic Rent System

This is a static GitHub Pages version of the Excel rent workbook.

## Editable fields

- `D2` -> Month date input at the top
- `H2:H5` -> Today's reading inputs for unit-based tenants
- `J7` -> Actual light bill input for Somnath More

## Converted VBA logic

When any `H2:H5` value is changed, the old value is copied to the corresponding `G` cell and the new value remains in `H`.

## How to publish on GitHub Pages

1. Create a new GitHub repository.
2. Upload `index.html`, `style.css`, and `script.js` to the repository root.
3. Go to **Settings > Pages**.
4. Select **Deploy from branch**.
5. Select branch `main` and folder `/root`.
6. Open the GitHub Pages URL after deployment.
