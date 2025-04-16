# Company Name Matching in Excel (VBA)

This VBA script helps compare two company names intelligently and outputs `Match` or `Not Match` based on similarity. It uses:

- ✅ **Levenshtein Distance**
- ✅ **Token Overlap**
- ✅ **Data Cleaning** (removes "Inc.", "LLC", punctuation, extra spaces, etc.)
- ✅ **Weighted Similarity Score**

---

## 📥 How to Add This Script to Excel

### 🔹 Step 1: Open Excel and Press `Alt + F11`
This opens the **VBA Editor**.

### 🔹 Step 2: Insert a New Module
- In the menu: click `Insert` > `Module`
- A blank code window will appear.

### 🔹 Step 3: Paste the VBA Script
- Copy all the code from `Module1.bas` (or this repo).
- Paste it into the module window.

---

## 🧪 How to Use in Excel Sheet

You can now use the function directly in Excel formulas:

```excel
=CheckCompanyMatch(A2, B2)
or
=CheckCompanyMatch(A2, B2, 70)
```

Where:

- `A2` is the first company name (e.g., from Apollo)
- `B2` is the second company name (e.g., from LinkedIn)
- Returns: `Match (85.00%)` or `Not Match (62.00%)` based on similarity

---

## ⚙️ Optional: Change Match Threshold

By default, the threshold is 70%. You can override it:

```excel
=CheckCompanyMatch(A2, B2, 80)
```

This uses 80% as the minimum similarity needed for a "Match".

---

## 💡 Example

| Apollo Name         | LinkedIn Name       | Result                 |
|---------------------|---------------------|-------------------------|
| 8x8                 | 8x8 UK              | ✅ Match (83.33%)       |
| Kearney             | A.T. Kearney        | ✅ Match (63.64%)       |
| Softura             | Altimetrik          | ❌ Not Match (10.00%)   |

---

## 📂 File List

- `Module1.bas` → Contains the full VBA script
- `README.md` → This instruction file

---

## 🧑‍💻 Author

Made with ❤️ by [Your Name]
