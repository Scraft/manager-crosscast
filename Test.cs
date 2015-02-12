using System;
using System.Linq;
using Manager;
using Manager.Model;
using Manager.Extensions;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Linq.Expressions;

namespace Test
{
    class Program
    {
        public struct PiExpense
        {
            public DateTime m_DateTime;
            public String m_Account;
            public String m_Contact;
            public String m_Description;
            public Boolean m_IsCashAcount;
            public String m_Category;
            public Decimal m_Amount;
            public Decimal m_Vat;
            public Guid? m_Currency;
            public Decimal m_AmountNativeCurrency;
            public Decimal m_VatNativeCurrency;
        };

        public struct PiExenseCategory
        {
            public Decimal m_Amount;
        }

        static Decimal GetExchangeRateFromCurrency(PersistentObjects objects, Guid? bankCurrency, Guid? transactionCurrency, DateTime time)
        {
            Guid? baseCurrency = GetBaseCurrency(objects);

            if (!bankCurrency.HasValue)
                bankCurrency = baseCurrency;

            if (!transactionCurrency.HasValue)
                transactionCurrency = baseCurrency;

            if (bankCurrency == transactionCurrency)
                return new Decimal(1.0f);

            Dictionary<DateTime, Decimal> currencyRates = new Dictionary<DateTime, Decimal>();

            foreach (ExchangeRates rates in objects.Values.OfType<ExchangeRates>())
            {
                foreach (ExchangeRates.ExchangeRate rate in rates.Rates)
                {
                    if (rate.Currency == bankCurrency.Value)
                    {
                        currencyRates[rates.Date] = rate.Rate.Value;
                    }
                }
            }

            DateTime result = currencyRates.Keys.Where(x => x <= time).Max();

            return currencyRates[result];
        }

        static Decimal GetExchangeRateFromAccountId(PersistentObjects objects, Guid? bankAccountId, Guid? transactionCurrency, DateTime time)
        {
            Guid? currency = null;
            BankAccount bankAccount = objects[bankAccountId.Value] as BankAccount;
            if (bankAccount != null)
                currency = bankAccount.Currency;
            CashAccount cashAccount = objects[bankAccountId.Value] as CashAccount;
            if (cashAccount != null)
                currency = cashAccount.Currency;

            if (!currency.HasValue || currency == transactionCurrency)
                return new Decimal( 1.0f );

            return GetExchangeRateFromCurrency(objects, currency, transactionCurrency, time);
        }

        static Decimal RoundTowardZero(Decimal val, int decimalPlaces)
        {
            Decimal scaler = new Decimal(Math.Pow(10, decimalPlaces));

            val *= scaler;

            Decimal frac = val - Math.Truncate(val);
            if (frac == new Decimal( 0.5 ) || frac == new Decimal( -0.5 ) )
            {
                return Math.Truncate(val) / scaler;
            }
            return Math.Round(val) / scaler;
        }

        public struct AccountDetails
        {
            public Guid? m_Account;
            public String m_Name;
            public Boolean m_IsCashAccount;
            public Boolean m_Valid;
            public Guid? m_Currency;
        }

        static AccountDetails GetAccountDetails(Guid? accountId, Manager.PersistentObjects objects)
        {
            AccountDetails accountDetails = new AccountDetails( );
            accountDetails.m_Account = accountId;
            accountDetails.m_Valid = false;

            if (accountId.HasValue)
            {
                BankAccount b = objects[accountId.Value] as BankAccount;
                CashAccount c = objects[accountId.Value] as CashAccount;
                if (b != null)
                {
                    accountDetails.m_Name = b.Name;
                    accountDetails.m_IsCashAccount = false;
                    accountDetails.m_Currency = b.Currency;
                    accountDetails.m_Valid = true;
                }
                else if (c != null)
                {
                    accountDetails.m_Name = c.Name;
                    accountDetails.m_IsCashAccount = true;
                    accountDetails.m_Currency = c.Currency;
                    accountDetails.m_Valid = true;
                }
            }

            return accountDetails;
        }

        static Guid? GetBaseCurrency(PersistentObjects objects)
        {
            List<BaseCurrency> baseCurrency = objects.Values.OfType<BaseCurrency>().ToList();
            Debug.Assert(baseCurrency.Count == 1, "Expected there to be a single base currency");
            return baseCurrency[0].Currency;
        }

        static void WorkOutAmountAndVat(Guid? taxCode, Guid? accountId, Guid? currencyId, DateTime date, Decimal nativeAmount, Boolean negate, PersistentObjects objects, Objects nonPersistentObjects, out Decimal amount, out Decimal vat, out Decimal amountNativeCurrency, out Decimal vatNativeCurrency)
        {
            Decimal taxRate = new Decimal(100.0f);

            if (taxCode.HasValue)
            {
                Manager.Query.TaxCodes.Item[] items = Manager.Query.TaxCodes.GetTaxCodes(nonPersistentObjects, taxCode.Value);
                foreach (Manager.Query.TaxCodes.Item item in items)
                {
                    foreach (Manager.Query.TaxCodes.Item.Component c in item.Components)
                    {
                        taxRate += c.Rate;
                    }
                }
            }

            // We want tax rate as a scaler (not a percentage).
            taxRate /= new Decimal(100.0f);

            // Amount in 'correct' currency, where it is the same currency as the account.
            {
                Decimal exchangeRate = GetExchangeRateFromAccountId(objects, accountId, currencyId, date);
                amount = (nativeAmount / taxRate) / exchangeRate;
                amount = RoundTowardZero(amount, 2);

                vat = (nativeAmount / exchangeRate) - amount;
                vat = RoundTowardZero(vat, 2);
            }

            // And the amount in native currency.
            {
                Guid? baseCurrency = GetBaseCurrency(objects);

                Decimal exchangeRate = GetExchangeRateFromCurrency(objects, currencyId, baseCurrency, date);
                amountNativeCurrency = (nativeAmount / taxRate) / exchangeRate;
                amountNativeCurrency = RoundTowardZero(amountNativeCurrency, 2);

                vatNativeCurrency = (nativeAmount / exchangeRate) - amountNativeCurrency;
                vatNativeCurrency = RoundTowardZero(vatNativeCurrency, 2);
            }

            if (negate)
            {
                amount = amount * -1;
                vat = vat * -1;
                amountNativeCurrency = amountNativeCurrency * -1;
                vatNativeCurrency = vatNativeCurrency * -1;
            }
        }

        public static String SplitByCasing(String input)
        {
            string output = "";

            foreach (char letter in input)
            {
                if (Char.IsUpper(letter) && output.Length > 0)
                    output += " " + letter;
                else
                    output += letter;
            }

            return output;
        }

        static String GetCategoryFromChartOfAccounts(Guid? accountId, out Boolean isFromChartOfAccounts )
        {
            isFromChartOfAccounts = false;

            Type type = typeof(Manager.Model.ChartOfAccounts);
            FieldInfo[] fields = type.GetFields();
            foreach (FieldInfo field in fields)
            {
                if (field.FieldType == typeof(Guid))
                {
                    if (accountId == (Guid)field.GetValue(null))
                    {
                        isFromChartOfAccounts = true;
                        return SplitByCasing(field.Name);
                    }
                }
            }

            return null;
        }

        static String GetCategoryFromAccountId(Guid? accountId, out Boolean isFromChartOfAccounts, Manager.PersistentObjects objects)
        {
            isFromChartOfAccounts = false;

            try
            {
                Manager.Model.Object obj = objects[accountId.Value];

                GeneralLedgerAccount g = obj as GeneralLedgerAccount;
                if (g != null)
                    return g.Name;
            }
            catch (System.Collections.Generic.KeyNotFoundException)
            {
            }

            return GetCategoryFromChartOfAccounts(accountId, out isFromChartOfAccounts);
        }

        static List<PiExpense> ProcessTransactionLines(Manager.Model.TransactionLine[] lines, AccountDetails accountDetails, DateTime date, String contact, String description, Boolean negate, Manager.PersistentObjects objects, Manager.Objects nonPersistentObjects)
        {
            List<PiExpense> expenses = new List<PiExpense>( );

            if (description == "Office furniture for new starters")
            {
                int a = 5;
            }

            foreach (TransactionLine l in lines)
            {
                PiExpense e = new PiExpense();
                e.m_Currency = accountDetails.m_Currency;

                bool valid = false;

                Guid? account = null;

                if (l.Amount == new Decimal(780.0f))
                {
                    int a = 5;
                }

                Guid? taxCode = null;

                if ( !account.HasValue && !taxCode.HasValue && l.PurchaseInvoice.HasValue)
                {
                    PurchaseInvoice invoice = objects[l.PurchaseInvoice.Value] as PurchaseInvoice;
                    Debug.Assert(invoice.Lines.Count() == 1, "Too many lines");
                    account = invoice.Lines[0].Account;
                    taxCode = invoice.Lines[0].TaxCode;
                }

                if (!account.HasValue && !taxCode.HasValue && l.SalesInvoice.HasValue)
                {
                    try
                    {
                        SalesInvoice invoice = objects[l.SalesInvoice.Value] as SalesInvoice;
                        Debug.Assert(invoice.Lines.Count() >= 1, "Not enough lines");
                        account = invoice.Lines[0].Account;
                        taxCode = invoice.Lines[0].TaxCode;
                        foreach (TransactionLine line in invoice.Lines)
                        {
                            Debug.Assert(line.Account == account, "Sales invoice with more than one category");
                            Debug.Assert(line.TaxCode == taxCode, "Sales invoice with more than one tax code");
                        }
                    }
                    catch (System.Collections.Generic.KeyNotFoundException)
                    {

                    }
                }

                if (!account.HasValue && !taxCode.HasValue && l.Account.HasValue)
                {
                    account = l.Account;
                }

                Boolean isFromChartOfAccounts = false;
                String category = GetCategoryFromAccountId(account, out isFromChartOfAccounts, objects);

                if (isFromChartOfAccounts)
                {
                    if (l.Member.HasValue)
                    {
                        CapitalAccount member = objects[l.Member.Value] as CapitalAccount;
                        category += "/";
                        category += member.Name;
                    }
                    if (l.MemberAccount.HasValue)
                    {
                        CapitalSubaccount memberAccount = objects[l.MemberAccount.Value] as CapitalSubaccount;
                        category += "/";
                        category += memberAccount.Name;
                    }
                    if (l.Employee.HasValue)
                    {
                        Employee employee = objects[l.Employee.Value] as Employee;
                        category += "/";
                        category += employee.Name;
                    }
                }

                if (account.HasValue)
                {
                    try
                    {
                        String bankAccountName = "<unknown>";

                        if (accountDetails.m_Valid)
                        {
                            bankAccountName = accountDetails.m_Name;
                            e.m_Account = accountDetails.m_Name;
                            e.m_IsCashAcount = accountDetails.m_IsCashAccount;
                        }

                        if (l.TaxCode.HasValue)
                            taxCode = l.TaxCode;

                        if (category != null)
                        {
                            valid = true;
                            e.m_Category = category;
                            e.m_DateTime = date;

                            if (e.m_Category == "Currency conversion")
                            {
                                int a = 5;
                            }

                            WorkOutAmountAndVat(taxCode, accountDetails.m_Account, accountDetails.m_Currency, date, l.Amount, negate, objects, nonPersistentObjects, out e.m_Amount, out e.m_Vat, out e.m_AmountNativeCurrency, out e.m_VatNativeCurrency);

                            if (category == "Bank charges")
                                Console.WriteLine(" {0} : {1} : {2} : {3} : {4}", bankAccountName, category, l.Description, contact, e.m_Amount);

                            e.m_Contact = contact;
                            e.m_Description = description;
                            expenses.Add(e);
                        }
                        else
                        {
                            ControlAccount c = objects[account.Value] as ControlAccount;

                            valid = true;
                            e.m_Category = c.Name;
                            if (e.m_Category == "")
                            {
                                e.m_Category = GetCategoryFromAccountId(account.Value, out isFromChartOfAccounts, objects);  
                            }                            
                            e.m_DateTime = date;

                            if (c.TaxCode.HasValue)
                                taxCode = c.TaxCode;

                            WorkOutAmountAndVat(taxCode, accountDetails.m_Account, accountDetails.m_Currency, date, l.Amount, true, objects, nonPersistentObjects, out e.m_Amount, out e.m_Vat, out e.m_AmountNativeCurrency, out e.m_VatNativeCurrency);

                            if (e.m_Category == null || e.m_Category == "")
                            {
                                int a = 5;
                            }
                            e.m_Contact = contact;
                            e.m_Description = description;
                            expenses.Add(e);
                        }
                    }
                    catch (System.Collections.Generic.KeyNotFoundException)
                    {

                    }
                }

                if (!valid)
                {
                    int a = 5;
                    //Console.WriteLine(" {0} : {1}", l.Description, l.Amount);
                }
            }

            return expenses;
        }

        static void Main()
        {
            // Open or create a file
            //var dicts = new PersistentDictionary("C:\\Projects\\Manager\\ManagerExporter\\bin\\shared\\Paw Print Games Ltd.manager");
            var objects = new PersistentObjects("C:\\Projects\\Manager-CrossCast\\bin\\shared\\Paw Print Games Ltd.manager");
            var nonPersistentObjects = new Objects("C:\\Projects\\Manager-CrossCast\\bin\\shared\\Objects");
            foreach (Guid g in objects.Keys)
            {
                nonPersistentObjects.Put(g, objects[g]);
            }

            
            foreach (JournalEntry j in objects.Values.OfType<JournalEntry>())
            {
                Console.WriteLine(j.Narration);
                foreach (TransactionLine l in j.Lines)
                {
                    Console.WriteLine(" {0} : {1}", l.Description, l.Amount);
                }
            }


            foreach (Transfer e in objects.Values.OfType<Transfer>())
            {
                BankAccount creditAccount = objects[e.CreditAccount.Value] as BankAccount;
                BankAccount debitAccount = objects[e.DebitAccount.Value] as BankAccount;
                if (debitAccount != null )
                {
                    Console.WriteLine("{0} {1} => {2}", e.Amount, creditAccount.Name, debitAccount.Name);
                }
                else
                {
                    CashAccount cashAccount = objects[e.DebitAccount.Value] as CashAccount;
                    if (cashAccount != null)
                    {
                        Console.WriteLine("{0} {1} => {2}", e.Amount, creditAccount.Name, cashAccount.Name);
                    }
                }
            }

            foreach (Transaction t in objects.Values.OfType<Transaction>())
            {
                Console.WriteLine("{0} : {1} : {2} : {3} : {4} : {5} : {6} : {7} : {8} ", t.Date.ToString(), t.Reference, t.Description, t.Contact, t.AccountAmount, t.BaseAmount, t.BaseTaxAmount, t.CurrencyConversion, t.TransactionAmount);
            }

            DateTime startDate = new DateTime(2014, 1, 7);
            DateTime endDate = new DateTime(2015, 1, 6);

            List<PiExpense> expenses = new List<PiExpense>();
            Dictionary<String, PiExenseCategory> expenseCategories = new Dictionary<String, PiExenseCategory>();

            foreach (Receipt r in objects.Values.OfType<Receipt>())
            {
                if (r.Date < startDate || r.Date > endDate)
                {
                    continue;
                }

                //Console.WriteLine("{0}", r.Lines[0].Amount);
                expenses.AddRange(ProcessTransactionLines(r.Lines, GetAccountDetails(r.DebitAccount, objects), r.Date, r.Payer, r.Description, false, objects, nonPersistentObjects));
            }

            foreach (Payment p in objects.Values.OfType<Payment>())
            {
                if (p.Date < startDate || p.Date > endDate)
                {
                    continue;
                }

                //Console.WriteLine("{0} : {1} : {2} : {3}", p.Date.ToString(), p.Payee, p.Reference, p.Description);
                expenses.AddRange( ProcessTransactionLines(p.Lines, GetAccountDetails(p.CreditAccount, objects), p.Date, p.Payee, p.Description, true, objects, nonPersistentObjects) );
            }

            // Build up categories from each expense.
            foreach (PiExpense expense in expenses)
            {
                if (! expenseCategories.ContainsKey(expense.m_Category) )
                {
                    PiExenseCategory category = new PiExenseCategory( );
                    category.m_Amount = 0;
                    expenseCategories.Add( expense.m_Category, category );
                }

                PiExenseCategory cat;
                if (expenseCategories.TryGetValue(expense.m_Category, out cat))
                {
                    cat.m_Amount += expense.m_Amount;
                    expenseCategories[expense.m_Category] = cat;
                }
            }

            // Sort by date (ascending).
            expenses.Sort((x, y) => x.m_DateTime.CompareTo(y.m_DateTime));

            // Get list of unique bank money accounts.
            List<String> accountTypes = expenses.Select(o => o.m_Account).Distinct().ToList( );

            // Sort account types alphabetically.
            accountTypes.Sort();

            // Get list of unique categories.
            List<String> categories = expenses.Select(o => o.m_Category).Distinct().ToList();

            // Sort category types alphabetically.
            categories.Sort( );

            System.Text.StringBuilder csv = new System.Text.StringBuilder();
            String newLine = "Date\tContact\tDescription\tAmount\t";
            foreach (String accountType in accountTypes)
            {
                newLine += accountType;
                newLine += "\t";
            }
            newLine += "VAT\tCost\t";
            foreach (String category in categories)
            {
                newLine += category;
                newLine += "\t";
            }
            newLine += string.Format("{0}", Environment.NewLine);

            csv.Append(newLine);

            foreach (PiExpense e in expenses)
            {
                Decimal gross = e.m_Amount + e.m_Vat;
                Decimal grossNative = e.m_AmountNativeCurrency + e.m_VatNativeCurrency;
                newLine = string.Format("{0}\t{1}\t{2}\t{3}", e.m_DateTime.ToShortDateString(), e.m_Contact, e.m_Description, grossNative);
                foreach (String accountType in accountTypes)
                {
                    newLine += "\t";
                    if (accountType == e.m_Account)
                        newLine += gross;
                }
                newLine += string.Format("\t{0}\t{1}", e.m_VatNativeCurrency != new Decimal(0) ? e.m_VatNativeCurrency.ToString() : "", e.m_AmountNativeCurrency);
                foreach (String category in categories)
                {
                    newLine += "\t";
                    if (category == e.m_Category)
                        newLine += e.m_AmountNativeCurrency;
                }
                newLine += string.Format("\t{0}", Environment.NewLine);
                csv.Append(newLine);
            }

            System.IO.File.WriteAllText("C:\\Projects\\Manager-CrossCast\\bin\\shared\\CrossCast.tsv", csv.ToString());
        }
    }
}