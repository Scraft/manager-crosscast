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
        static String kCurrencyGainsLosses = "Currency gains/losses";

        public struct PiAccountTransaction
        {
            public Boolean m_IsCashAcount;
            public String m_AccountName;
            public Guid? m_Currency;
            public Decimal m_Amount;
            public Decimal m_Vat;
        };
        public struct PiExpense
        {
            public Boolean m_InRequestedPeriod;
            public DateTime m_DateTime;
            public String m_Contact;
            public String m_Description;
            public String m_Category;
            public Decimal m_AmountNativeCurrency;
            public Decimal m_VatNativeCurrency;
            public Decimal m_AmountNativeCurrencyAtInvoiceDate;
            public Decimal m_VatNativeCurrencyAtInvoiceDate;
            public PiAccountTransaction m_SourceTransaction;
            public PiAccountTransaction m_DestinationTransaction;
            public Boolean m_IsTransfer;
        };

        public struct PiExenseCategory
        {
            public Decimal m_Amount;
        }

        static void Swap<T>(ref T lhs, ref T rhs)
        {
            T temp;
            temp = lhs;
            lhs = rhs;
            rhs = temp;
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

            Debug.Assert(bankCurrency == baseCurrency || transactionCurrency == baseCurrency, "One of the currencies must be the base currency");
            Boolean swapped = false;
            if (bankCurrency == baseCurrency)
            {
                Swap(ref bankCurrency, ref transactionCurrency);
                swapped = true;
            }

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

            Decimal resultRate = currencyRates[result];

            if (swapped)
                resultRate = new Decimal(1.0f) / resultRate;

            return resultRate;
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

        static Decimal WorkOutAmountInBaseCurrency(Guid? currencyId, DateTime date, Decimal amount, PersistentObjects objects)
        {
            return amount * GetExchangeRateFromCurrency(objects, GetBaseCurrency(objects), currencyId, date);
        }

        static void WorkOutAmountAndVat(Guid? taxCode, Guid? accountId, Guid? currencyId, DateTime date, DateTime invoiceDate, Guid? invoiceCurrency, Decimal nativeAmount, Boolean negate, PersistentObjects objects, Objects nonPersistentObjects, out Decimal amount, out Decimal vat, out Decimal amountNativeCurrency, out Decimal vatNativeCurrency, out Decimal amountNativeCurrencyAtInvoiceDate, out Decimal vatNativeCurrencyAtInvoiceDate)
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

            // And the amount in native currency on invoice date.
            // Note, if a US invoice is raised, and then paid into a UK bank account, we need to work out how much the GBP
            // money received would have been worth at invoice date.
            {
                // Get amount in invoice currency.
                {
                    Decimal newExchangeRate = GetExchangeRateFromCurrency(objects, currencyId, invoiceCurrency, date);
                    Decimal test = nativeAmount / newExchangeRate;
                    Decimal oldExchangeRate = GetExchangeRateFromCurrency(objects, currencyId, invoiceCurrency, invoiceDate);
                    test = test * oldExchangeRate;

                    nativeAmount = RoundTowardZero(test, 2);
                }

                {
                    Guid? baseCurrency = GetBaseCurrency(objects);

                    Decimal exchangeRate = GetExchangeRateFromCurrency(objects, currencyId, baseCurrency, invoiceDate);
                    amountNativeCurrencyAtInvoiceDate = (nativeAmount / taxRate) / exchangeRate;
                    amountNativeCurrencyAtInvoiceDate = RoundTowardZero(amountNativeCurrencyAtInvoiceDate, 2);

                    vatNativeCurrencyAtInvoiceDate = (nativeAmount / exchangeRate) - amountNativeCurrencyAtInvoiceDate;
                    vatNativeCurrencyAtInvoiceDate = RoundTowardZero(vatNativeCurrencyAtInvoiceDate, 2);
                }
            }

            if (negate)
            {
                amount = amount * -1;
                vat = vat * -1;
                amountNativeCurrency = amountNativeCurrency * -1;
                vatNativeCurrency = vatNativeCurrency * -1;
                amountNativeCurrencyAtInvoiceDate = amountNativeCurrencyAtInvoiceDate * -1;
                vatNativeCurrencyAtInvoiceDate = vatNativeCurrencyAtInvoiceDate * -1;
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
            if (!accountId.HasValue)
            {
                isFromChartOfAccounts = false;
                return null;
            }

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

        static List<PiExpense> ProcessTransactionLines(Manager.Model.TransactionLine[] lines, AccountDetails accountDetails, DateTime date, DateTime startDate, DateTime endDate, String contact, String description, Boolean negate, Manager.PersistentObjects objects, Manager.Objects nonPersistentObjects)
        {
            List<PiExpense> expenses = new List<PiExpense>( );

            foreach (TransactionLine l in lines)
            {
                DateTime invoiceDate = date;
                Guid? invoiceCurrency = GetBaseCurrency(objects);

                PiExpense e = new PiExpense();
                e.m_InRequestedPeriod = date >= startDate && date <= endDate;
                e.m_SourceTransaction.m_Currency = accountDetails.m_Currency;

                bool valid = false;

                Guid? account = null;
                Guid? taxCode = null;

                if ( !account.HasValue && !taxCode.HasValue && l.PurchaseInvoice.HasValue)
                {
                    PurchaseInvoice invoice = objects[l.PurchaseInvoice.Value] as PurchaseInvoice;
                    Debug.Assert(invoice.Lines.Count() == 1, "Too many lines");
                    account = invoice.Lines[0].Account;
                    taxCode = invoice.Lines[0].TaxCode;
                    invoiceDate = invoice.IssueDate;
                }

                if (!account.HasValue && !taxCode.HasValue && l.SalesInvoice.HasValue)
                {
                    try
                    {
                        SalesInvoice invoice = objects[l.SalesInvoice.Value] as SalesInvoice;
                        Debug.Assert(invoice.Lines.Count() >= 1, "Not enough lines");
                        account = invoice.Lines[0].Account;
                        taxCode = invoice.Lines[0].TaxCode;
                        invoiceDate = invoice.IssueDate;

                        if (invoice.To.HasValue)
                        {
                            Customer customer = objects[invoice.To.Value] as Customer;
                            invoiceCurrency = customer.Currency;
                        }

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
                    if (l.FixedAsset.HasValue)
                    {
                        FixedAsset fixedAsset = objects[l.FixedAsset.Value] as FixedAsset;
                        category += "/";
                        category += fixedAsset.Name;
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
                            e.m_SourceTransaction.m_AccountName = accountDetails.m_Name;
                            e.m_SourceTransaction.m_IsCashAcount = accountDetails.m_IsCashAccount;
                        }

                        if (l.TaxCode.HasValue)
                            taxCode = l.TaxCode;

                        if (category != null)
                        {
                            valid = true;
                            e.m_Category = category;
                            e.m_DateTime = date;

                            if (!accountDetails.m_Account.HasValue)
                            {
                                // Journal entry or something else that doesn't change any bank/cash accounts. Assume it is always native currency with no VAT.
                                e.m_SourceTransaction.m_Amount = 0;
                                e.m_SourceTransaction.m_Vat = 0;
                                e.m_AmountNativeCurrency = (l.Credit.HasValue ? l.Credit.Value : 0) - (l.Debit.HasValue ? l.Debit.Value : 0);
                                e.m_VatNativeCurrency = 0;
                                e.m_AmountNativeCurrencyAtInvoiceDate = e.m_AmountNativeCurrency;
                                e.m_VatNativeCurrencyAtInvoiceDate = 0;
                            }
                            else
                            {
                                WorkOutAmountAndVat(taxCode, accountDetails.m_Account, accountDetails.m_Currency, date, invoiceDate, invoiceCurrency, l.Amount, negate, objects, nonPersistentObjects, out e.m_SourceTransaction.m_Amount, out e.m_SourceTransaction.m_Vat, out e.m_AmountNativeCurrency, out e.m_VatNativeCurrency, out e.m_AmountNativeCurrencyAtInvoiceDate, out e.m_VatNativeCurrencyAtInvoiceDate);
                            }

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

                            WorkOutAmountAndVat(taxCode, accountDetails.m_Account, accountDetails.m_Currency, date, invoiceDate, invoiceCurrency, l.Amount, true, objects, nonPersistentObjects, out e.m_SourceTransaction.m_Amount, out e.m_SourceTransaction.m_Vat, out e.m_AmountNativeCurrency, out e.m_VatNativeCurrency, out e.m_AmountNativeCurrencyAtInvoiceDate, out e.m_VatNativeCurrencyAtInvoiceDate);

                            e.m_Contact = contact;
                            e.m_Description = description;
                            expenses.Add(e);
                        }
                    }
                    catch (System.Collections.Generic.KeyNotFoundException)
                    {

                    }
                }

                // Any transactions not yet categorised will appear here.
                //Debug.Assert(valid, "Could not process");
            }

            return expenses;
        }

        static void Main()
        {
            // Open or create a file
            //var dicts = new PersistentDictionary("C:\\Projects\\Manager\\ManagerExporter\\bin\\shared\\Paw Print Games Ltd.manager");
            var objects = new PersistentObjects("C:\\Projects\\Manager-CrossCast\\bin\\shared\\Paw Print Games Ltd.manager");
            String nonPersistentObjectPath = "C:\\Projects\\Manager-CrossCast\\bin\\shared\\Objects";
            if (System.IO.Directory.Exists(nonPersistentObjectPath))
            {
                System.IO.Directory.Delete(nonPersistentObjectPath, true);
            }
            var nonPersistentObjects = new Objects(nonPersistentObjectPath);
            foreach (Guid g in objects.Keys)
            {
                nonPersistentObjects.Put(g, objects[g]);
            }

            DateTime startDate = new DateTime(2014, 1, 7);
            DateTime endDate = new DateTime(2015, 1, 6);

            List<PiExpense> expenses = new List<PiExpense>();
            Dictionary<String, PiExenseCategory> expenseCategories = new Dictionary<String, PiExenseCategory>();


            foreach (Transfer t in objects.Values.OfType<Transfer>())
            {
                AccountDetails creditAccountDetails = GetAccountDetails(t.CreditAccount.Value, objects);
                AccountDetails debitAccountDetails = GetAccountDetails(t.DebitAccount.Value, objects);

                BankAccount creditAccount = objects[t.CreditAccount.Value] as BankAccount;
                BankAccount debitAccount = objects[t.DebitAccount.Value] as BankAccount;

                PiExpense e;
                e.m_InRequestedPeriod = t.Date >= startDate && t.Date <= endDate;
                e.m_DateTime = t.Date;
                e.m_Contact = "Transfer";
                e.m_Description = t.Description;
                e.m_Category = "Transfer";

                e.m_SourceTransaction.m_IsCashAcount = creditAccountDetails.m_IsCashAccount;
                e.m_SourceTransaction.m_AccountName = creditAccountDetails.m_Name;
                e.m_SourceTransaction.m_Currency = creditAccountDetails.m_Currency;
                WorkOutAmountAndVat(null, t.CreditAccount, GetBaseCurrency(objects), t.Date, t.Date, GetBaseCurrency(objects), t.Amount, true, objects, nonPersistentObjects, out e.m_SourceTransaction.m_Amount, out e.m_SourceTransaction.m_Vat, out e.m_AmountNativeCurrency, out e.m_VatNativeCurrency, out e.m_AmountNativeCurrencyAtInvoiceDate, out e.m_VatNativeCurrencyAtInvoiceDate);

                e.m_DestinationTransaction.m_IsCashAcount = debitAccountDetails.m_IsCashAccount;
                e.m_DestinationTransaction.m_AccountName = debitAccountDetails.m_Name;
                e.m_DestinationTransaction.m_Currency = debitAccountDetails.m_Currency;
                WorkOutAmountAndVat(null, t.CreditAccount, GetBaseCurrency(objects), t.Date, t.Date, GetBaseCurrency(objects), t.Amount, false, objects, nonPersistentObjects, out e.m_DestinationTransaction.m_Amount, out e.m_DestinationTransaction.m_Vat, out e.m_AmountNativeCurrency, out e.m_VatNativeCurrency, out e.m_AmountNativeCurrencyAtInvoiceDate, out e.m_VatNativeCurrencyAtInvoiceDate);

                e.m_IsTransfer = true;

                expenses.Add(e);
            }

            if (true)
            {
                foreach (JournalEntry j in objects.Values.OfType<JournalEntry>())
                {
                    AccountDetails accountDetails = new AccountDetails();
                    expenses.AddRange(ProcessTransactionLines(j.Lines, accountDetails, j.Date, startDate, endDate, "Journal entry", j.Narration, false, objects, nonPersistentObjects));
                }
            }

            foreach (Receipt r in objects.Values.OfType<Receipt>())
            {
                expenses.AddRange(ProcessTransactionLines(r.Lines, GetAccountDetails(r.DebitAccount, objects), r.Date, startDate, endDate, r.Payer, r.Description, false, objects, nonPersistentObjects));
            }

            foreach (Payment p in objects.Values.OfType<Payment>())
            {
                expenses.AddRange(ProcessTransactionLines(p.Lines, GetAccountDetails(p.CreditAccount, objects), p.Date, startDate, endDate, p.Payee, p.Description, true, objects, nonPersistentObjects));
            }

            // Get bank account opening balances.
            Dictionary<String, Dictionary<DateTime, Decimal>> bankAccountToBalance = new Dictionary<String, Dictionary<DateTime, Decimal>>();

            foreach (BankAccount b in objects.Values.OfType<BankAccount>())
            {
                bankAccountToBalance[b.Name] = new Dictionary<DateTime, Decimal>(); 
                bankAccountToBalance[b.Name][new DateTime(0)] = b.StartingBalance;
            }

            // Get cash account opening balances.
            Dictionary<String, Dictionary<DateTime, Decimal>> cashAccountToBalance = new Dictionary<String, Dictionary<DateTime, Decimal>>();

            foreach (CashAccount c in objects.Values.OfType<CashAccount>())
            {
                cashAccountToBalance[c.Name] = new Dictionary<DateTime, Decimal>(); 
                cashAccountToBalance[c.Name][new DateTime(0)] = c.StartingBalance;
            }

            // Go through (date sorted) expenses and add to bank balances.
            List<PiExpense> dateSortedExpenses = expenses;
            dateSortedExpenses.Sort((x, y) => x.m_DateTime.CompareTo(y.m_DateTime)); 
            foreach ( PiExpense e in dateSortedExpenses )
            {
                if (e.m_IsTransfer)
                {
                    int a = 5;
                    ++a;
                }

                if (e.m_SourceTransaction.m_AccountName != null)
                {
                    if (e.m_SourceTransaction.m_IsCashAcount)
                    {
                        Decimal balance;
                        if ( cashAccountToBalance[e.m_SourceTransaction.m_AccountName].ContainsKey(e.m_DateTime) )
                            balance = cashAccountToBalance[e.m_SourceTransaction.m_AccountName][e.m_DateTime];
                        else
                            balance = cashAccountToBalance[e.m_SourceTransaction.m_AccountName].Last().Value;
                        cashAccountToBalance[e.m_SourceTransaction.m_AccountName][e.m_DateTime] = balance + e.m_SourceTransaction.m_Amount;
                    }
                    else
                    {
                        Decimal balance;
                        if (bankAccountToBalance[e.m_SourceTransaction.m_AccountName].ContainsKey(e.m_DateTime))
                            balance = bankAccountToBalance[e.m_SourceTransaction.m_AccountName][e.m_DateTime];
                        else
                            balance = bankAccountToBalance[e.m_SourceTransaction.m_AccountName].Last().Value;
                        bankAccountToBalance[e.m_SourceTransaction.m_AccountName][e.m_DateTime] = balance + e.m_SourceTransaction.m_Amount;
                    }
                }

                if (e.m_DestinationTransaction.m_AccountName != null)
                {
                    if (e.m_DestinationTransaction.m_IsCashAcount)
                    {
                        Decimal balance;
                        if (cashAccountToBalance[e.m_DestinationTransaction.m_AccountName].ContainsKey(e.m_DateTime))
                            balance = cashAccountToBalance[e.m_DestinationTransaction.m_AccountName][e.m_DateTime];
                        else
                            balance = cashAccountToBalance[e.m_DestinationTransaction.m_AccountName].Last().Value;
                        cashAccountToBalance[e.m_DestinationTransaction.m_AccountName][e.m_DateTime] = balance + e.m_DestinationTransaction.m_Amount;
                    }
                    else
                    {
                        Decimal balance;
                        if (bankAccountToBalance[e.m_DestinationTransaction.m_AccountName].ContainsKey(e.m_DateTime))
                            balance = bankAccountToBalance[e.m_DestinationTransaction.m_AccountName][e.m_DateTime];
                        else
                            balance = bankAccountToBalance[e.m_DestinationTransaction.m_AccountName].Last().Value;
                        bankAccountToBalance[e.m_DestinationTransaction.m_AccountName][e.m_DateTime] = balance + e.m_DestinationTransaction.m_Amount;
                    }
                }

            }

            // Currency gains and losses.
            Guid? baseCurrency = GetBaseCurrency(objects);
            Dictionary<Guid, DateTime?> currencyGuidToLastExchangeRate = new Dictionary<Guid,DateTime?>();
            List<ExchangeRates> exchangeRatesDateSorted = objects.Values.OfType<ExchangeRates>().ToList();
            exchangeRatesDateSorted.Sort((x, y) => x.Date.CompareTo(y.Date));
            foreach (ExchangeRates rates in exchangeRatesDateSorted )
            {
                foreach (ExchangeRates.ExchangeRate rate in rates.Rates)
                {
                    if (rate.Currency != baseCurrency.Value)
                    {
                        if (!currencyGuidToLastExchangeRate.Keys.Contains(rate.Currency))
                            currencyGuidToLastExchangeRate[rate.Currency] = rates.Date;

                        List<PiExpense> tempExpenses = new List<PiExpense>();

                        // Go through any accounts (bank/cash) which are in the specified currency, and apply gain/loss.
                        foreach ( BankAccount b in objects.Values.OfType<BankAccount>())
                        {
                            if ( b.Currency == rate.Currency )
                            {
                                DateTime? lastExchangeRateDate = currencyGuidToLastExchangeRate[rate.Currency];
                                PiExpense e = CreateCurrencyGainLossExpense(rates, bankAccountToBalance, b.Name, b.Currency, ref lastExchangeRateDate, objects, startDate, endDate);
                                currencyGuidToLastExchangeRate[rate.Currency] = lastExchangeRateDate;
                                tempExpenses.Add(e);
                            }
                        }
                        foreach (CashAccount c in objects.Values.OfType<CashAccount>())
                        {
                            if (c.Currency == rate.Currency)
                            {
                                DateTime? lastExchangeRateDate = currencyGuidToLastExchangeRate[rate.Currency];
                                PiExpense e = CreateCurrencyGainLossExpense(rates, cashAccountToBalance, c.Name, c.Currency, ref lastExchangeRateDate, objects, startDate, endDate);
                                currencyGuidToLastExchangeRate[rate.Currency] = lastExchangeRateDate;
                                tempExpenses.Add(e);
                            }
                        }

                        // We want to combine each days activity into one.
                        if (tempExpenses.Count > 0)
                        {
                            PiExpense expense = tempExpenses[0];

                            bool first = true;
                            foreach (PiExpense e in tempExpenses)
                            {
                                if (first == true)
                                {
                                    first = false;
                                    continue;
                                }
                                expense.m_AmountNativeCurrency += e.m_AmountNativeCurrency;
                                expense.m_AmountNativeCurrencyAtInvoiceDate += e.m_AmountNativeCurrencyAtInvoiceDate;
                            }

                            if (expense.m_AmountNativeCurrency != expense.m_AmountNativeCurrencyAtInvoiceDate)
                            {
                                Decimal oldExchange = GetExchangeRateFromCurrency(objects, GetBaseCurrency(objects), rate.Currency, currencyGuidToLastExchangeRate[rate.Currency].Value);
                                Decimal newExchange = GetExchangeRateFromCurrency(objects, GetBaseCurrency(objects), rate.Currency, rates.Date);
                                expense.m_Description = String.Format("Currency rate change ({0} -> {1}) for {2}", RoundTowardZero( oldExchange, 4 ), RoundTowardZero( newExchange, 4 ), RoundTowardZero( expense.m_AmountNativeCurrencyAtInvoiceDate, 2 ) );

                                expense.m_AmountNativeCurrency = RoundTowardZero(expense.m_AmountNativeCurrency, 2);
                                expense.m_AmountNativeCurrencyAtInvoiceDate = RoundTowardZero(expense.m_AmountNativeCurrencyAtInvoiceDate, 2);
                                expenses.Add(expense);
                            }
                        }
                    }

                    currencyGuidToLastExchangeRate[rate.Currency] = rates.Date;
                }
            }

            // We only want to put expenses within out date range onto the cross cast.
            expenses = expenses.Where(e => e.m_InRequestedPeriod == true).ToList();

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
                    cat.m_Amount += expense.m_SourceTransaction.m_Amount;
                    expenseCategories[expense.m_Category] = cat;
                }
            }

            // Sort by date (ascending).
            expenses.Sort((x, y) => x.m_DateTime.CompareTo(y.m_DateTime));

            // Get list of unique bank money accounts.
            List<String> accountTypes = expenses.Select(o => o.m_SourceTransaction.m_AccountName).Distinct().Except(new List<String> {null}).ToList();

            // Sort account types alphabetically.
            accountTypes.Sort();

            // Get list of unique categories.
            List<String> categories = expenses.Select(o => o.m_Category).Distinct().ToList();

            // Also want a category for currency gains/losses (if any native currencies do not match native currency at invoice date).
            if (!categories.Contains(kCurrencyGainsLosses) && expenses.Select(o => o.m_AmountNativeCurrency != o.m_AmountNativeCurrencyAtInvoiceDate).Contains(true))
            {
                categories.Add(kCurrencyGainsLosses);
            }

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
                Decimal sourceGross = e.m_SourceTransaction.m_Amount + e.m_SourceTransaction.m_Vat;
                Decimal destinationGross = e.m_DestinationTransaction.m_Amount + e.m_DestinationTransaction.m_Vat;
                Decimal grossNative = e.m_AmountNativeCurrency + e.m_VatNativeCurrency;
                Decimal nativeAmount = e.m_AmountNativeCurrency;
                Decimal nativeAmountAtInvoiceDate = e.m_AmountNativeCurrencyAtInvoiceDate;
                Decimal nativeCurrencyGain = e.m_AmountNativeCurrency - e.m_AmountNativeCurrencyAtInvoiceDate;

                // If this is a currency gain, then we don't want to use any numbers except the currency gain value.
                if (e.m_Category.CompareTo(kCurrencyGainsLosses) == 0)
                {
                    sourceGross = 0;
                    destinationGross = 0;
                    grossNative = 0;
                    nativeAmount = 0;
                    nativeAmountAtInvoiceDate = 0;
                }

                if (e.m_IsTransfer)
                {
                    grossNative = new Decimal(0);
                    nativeAmount = new Decimal(0);
                    nativeAmountAtInvoiceDate = new Decimal(0);
                }
                newLine = string.Format("{0}\t{1}\t{2}\t{3}", e.m_DateTime.ToShortDateString(), e.m_Contact, e.m_Description, grossNative);
                foreach (String accountType in accountTypes)
                {
                    newLine += "\t";
                    if (accountType == e.m_SourceTransaction.m_AccountName)
                        newLine += sourceGross;
                    if (accountType == e.m_DestinationTransaction.m_AccountName)
                        newLine += destinationGross;
                }
                newLine += string.Format("\t{0}\t{1}", e.m_VatNativeCurrency != new Decimal(0) ? e.m_VatNativeCurrency.ToString() : "", nativeAmount);
                foreach (String category in categories)
                {
                    newLine += "\t";
                    if (category == kCurrencyGainsLosses && nativeCurrencyGain != new Decimal(0))
                        newLine += nativeCurrencyGain;
                    else if (category == e.m_Category)
                        newLine += nativeAmountAtInvoiceDate;
                }
                newLine += string.Format("\t{0}", Environment.NewLine);
                csv.Append(newLine);
            }

            System.IO.File.WriteAllText("C:\\Projects\\Manager-CrossCast\\bin\\shared\\CrossCast.tsv", csv.ToString());
        }

        private static PiExpense CreateCurrencyGainLossExpense(ExchangeRates rates, Dictionary<String, Dictionary<DateTime, Decimal>> bankAccountToBalance, String accountName, Guid? accountCurrency, ref DateTime? lastExchangeRateDate, Manager.PersistentObjects objects, DateTime startDate, DateTime endDate)
        {
            PiExpense e = new PiExpense();
            e.m_DateTime = rates.Date;
            e.m_Contact = kCurrencyGainsLosses;
            e.m_Description = kCurrencyGainsLosses;
            e.m_Category = kCurrencyGainsLosses;

            e.m_AmountNativeCurrency = 0;
            e.m_VatNativeCurrency = 0;
            e.m_AmountNativeCurrencyAtInvoiceDate = 0;
            e.m_VatNativeCurrencyAtInvoiceDate = 0;

            // We use one day in the past, as we don't want to include the specified days transactions.
            DateTime lookUpDate = rates.Date;
            lookUpDate = lookUpDate.AddDays(-1.0);

            // Get the amount of money in the bank at a given date... somehow!
            int ix = bankAccountToBalance[accountName].Keys.ToList().BinarySearch(lookUpDate);
            if (ix < 0)
                ix = ~ix - 1;

            Decimal foreignCurrencyAmount = bankAccountToBalance[accountName].Values.ToList()[ix];
            Decimal previousCurrencyAmount = WorkOutAmountInBaseCurrency(accountCurrency, lastExchangeRateDate.Value, foreignCurrencyAmount, objects);
            Decimal nativeCurrencyAmount = WorkOutAmountInBaseCurrency(accountCurrency, rates.Date, foreignCurrencyAmount, objects);
            Decimal gainLoss = nativeCurrencyAmount - previousCurrencyAmount;

            e.m_AmountNativeCurrency = nativeCurrencyAmount;
            e.m_AmountNativeCurrencyAtInvoiceDate = previousCurrencyAmount;

            e.m_InRequestedPeriod = e.m_DateTime >= startDate && e.m_DateTime <= endDate;
            e.m_IsTransfer = false;

            return e;
        }
    }
}