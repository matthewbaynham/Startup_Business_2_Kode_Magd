using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using PaymillWrapper;
//using Newtonsoft;
//using System.Net.Http.Formatting;
//using PaymillWrapper;

namespace KodeMagd.Payment
{
    class ClsPaymentTest
    {
        /*
         * Test all these code examples in this link
         * 
         https://paymill.codeplex.com/SourceControl/latest#UnitTest/Net/TestPayments.cs
         */


        /*
         * PaymillWrapper pmw = new PaymillWrapper();





        //Authentication
        //To authenticate at the Paymill Wrapper NET API, you need the private key of your test or live account.
        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";

        //Payments
        //The Payment object represents a payment with a credit card or via direct debit.
        //Create new Credit Card Payment

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        PaymentService paymentService = Paymill.GetService<PaymentService>();

        Payment payment = new Payment();
        payment.Token = "098f6bcd4621d373cade4e832627b4f6";
         
        Payment newPayment = paymentService.AddPayment(payment);

        //Create new Debit Payment

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        PaymentService paymentService = Paymill.GetService<PaymentService>();

        Payment payment = new Payment();
        payment.Type = Payment.TypePayment.DEBIT;
        payment.Code = "86055500";
        payment.Account = "1234512345";
        payment.Holder = "Max Mustermann";

        Payment newPayment = paymentService.AddPayment(payment);

        Payment Details

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        PaymentService paymentService = Paymill.GetService<PaymentService>();

        string paymentID = "pay_4c159fe95d3be503778a";
        Payment payment = paymentService.GetPayment(paymentID);

        //List Payments

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        PaymentService paymentService = Paymill.GetService<PaymentService>();

        List<Payment> lstPayments = paymentService.GetPayments();

        //List Payments with filter

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        PaymentService paymentService = Paymill.GetService<PaymentService>();

        Filter filter = new Filter();
        filter.Add("count", 5);
        filter.Add("offset", 41);

        List<Payment> lstPayments = paymentService.GetPayments(filter);

        //Remove Payment

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        PaymentService paymentService = Paymill.GetService<PaymentService>();

        string paymentID = "pay_640be2127169cea1d375";
        bool reply = paymentService.RemovePayment(paymentID);

        //Transactions
        //A transaction is the charging of a credit card or a direct debit. In this case you need a new transaction object with either a valid token, payment, client + payment or preauthorization.
        //Create new Transaction

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        TransactionService transactionService = Paymill.GetService<TransactionService>();

        Transaction transaction = new Transaction();
        transaction.Token = "098f6bcd4621d373cade4e832627b4f6";
        transaction.Amount = 3500;
        transaction.Currency = "EUR";
        transaction.Description = "Prueba desde API c#";

        Transaction newTransaction = transactionService.AddTransaction(transaction);

        //Transaction Details

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        TransactionService transactionService = Paymill.GetService<TransactionService>();

        string transactionID = "tran_9255ee9ad5a7f2999625";
        Transaction transaction = transactionService.GetTransaction(transactionID);

        //List Transactions

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        TransactionService transactionService = Paymill.GetService<TransactionService>();

        List<Transaction> lstTransactions = transactionService.GetTransactions();

        //List Transactions with filter

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        TransactionService transactionService = Paymill.GetService<TransactionService>();

        Filter filter = new Filter();
        filter.Add("count", 1);
        filter.Add("offset", 2);

        List<Transaction> lstTransactions = transactionService.GetTransactions(filter);

        //Refunds
        //Refunds are own objects with own calls for existing transactions. The refunded amount will be credited to the account of the client.
        //Refund Transaction

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        RefundService refundService = Paymill.GetService<RefundService>();

        Refund refund = new Refund();
        refund.Amount = 500;
        refund.Description = "Prueba desde API c#";
        refund.Transaction = new Transaction() { Id = "tran_a7c93a1e5b431b52c0f0" };

        Refund newRefund = refundService.AddRefund(refund);

        //Refund Details

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        RefundService refundService = Paymill.GetService<RefundService>();

        string refundID = "refund_53860aa0e514d4913aad";
        Refund refund = refundService.GetRefund(refundID);

        //List Refunds

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        RefundService refundService = Paymill.GetService<RefundService>();

        List<Refund> lstRefunds = refundService.GetRefunds();

        //List Refunds with filter

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        RefundService refundService = Paymill.GetService<RefundService>();

        Filter filter = new Filter();
        filter.Add("count", 5);

        List<Refund> lstRefunds = refundService.GetRefunds(filter);

        //Clients
        //The clients object is used to edit, delete, update clients as well as to permit refunds, subscriptions, insert credit card details for a client, edit client details and of course make transactions.
        //Create new Client

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        ClientService clientService = Paymill.GetService<ClientService>();

        Client c = new Client();
        c.Description = "Prueba API";
        c.Email = "javicantos22@hotmail.es";

        Client newClient = clientService.AddClient(c);

        //Client Details

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        ClientService clientService = Paymill.GetService<ClientService>();

        string clientID = "client_ad591663d69051d306a8";
        Client c = clientService.GetClient(clientID);

        //Update Client

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        ClientService clientService = Paymill.GetService<ClientService>();

        Client c = new Client();
        c.Description = "Javier";
        c.Email = "javicantos33@hotmail.es";
        c.Id = "client_bbe895116de80b6141fd";

        Client updatedClient = clientService.UpdateClient(c);

        //Remove Client

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        ClientService clientService = Paymill.GetService<ClientService>();

        string clientID = "client_180ad3d1042a1da4a0a0";
        bool reply = clientService.RemoveClient(clientID);

        //List Clients

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        ClientService clientService = Paymill.GetService<ClientService>();

        List<Client> lstClients = clientService.GetClients();

        //List Clients with filter

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        ClientService clientService = Paymill.GetService<ClientService>();

        Filter filter = new Filter();
        filter.Add("email", "jcantos@gmail.com"); OK

        List<Client> lstClients = clientService.GetClients(filter);

        //Offers
        //An offer is a recurring plan which a user can subscribe to. You can create different offers with different plan attributes e.g. a monthly or a yearly based paid offer/plan.
        //Create new Offer

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        OfferService offerService = Paymill.GetService<OfferService>();

        Offer offer = new Offer();
        offer.Amount = 1500;
        offer.Currency = "eur";
        offer.Interval = Offer.TypeInterval.MONTH;
        offer.Name = "Oferta 24";
        offer.Trial_Period_Days = 3;

        Offer newOffer = offerService.AddOffer(offer);

        //Offer Details

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        OfferService offerService = Paymill.GetService<OfferService>();

        string offerID = "offer_6eea405f83d4d3098604";
        Offer offer = offerService.GetOffer(offerID);

        //Update Offer

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        OfferService offerService = Paymill.GetService<OfferService>();

        Offer offer = new Offer();
        offer.Name = "Oferta 48";
        offer.Id = "offer_6eea405f83d4d3098604";

        Offer updatedOffer = offerService.UpdateOffer(offer);

        //Remove Offer

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        OfferService offerService = Paymill.GetService<OfferService>();

        string offerID = "offer_6eea405f83d4d3098604";
        bool reply = offerService.RemoveOffer(offerID);

        //List Offers

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        OfferService offerService = Paymill.GetService<OfferService>();

        List<Offer> lstOffers = offerService.GetOffers();

        //List Offers with filter

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        OfferService offerService = Paymill.GetService<OfferService>();

        Filter filter = new Filter();
        filter.Add("interval","month"); //OK
        filter.Add("amount", 495); //OK

        List<Offer> lstOffers = offerService.GetOffers(filter);

        //Subscriptions
        //Subscriptions allow you to charge recurring payments on a client’s credit card / to a client’s direct debit. A subscription connects a client to the offers-object.
        //Create new Subscription

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        SubscriptionService susbscriptionService = Paymill.GetService<SubscriptionService>();

        Subscription subscription = new Subscription();
        subscription.Client = new Client() { Id = "client_bbe895116de80b6141fd" };
        subscription.Offer = new Offer() { Id = "offer_32008ddd39954e71ed48" };
        subscription.Payment = new Payment() { Id = "pay_81ec02206e9b9c587513" };

        Subscription newSubscription = susbscriptionService.AddSubscription(subscription);

        //Subscription Details

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        SubscriptionService susbscriptionService = Paymill.GetService<SubscriptionService>();

        string subscriptionID = "sub_e77d3332e456674101ad";
        Subscription subscription = susbscriptionService.GetSubscription(subscriptionID);

        //Update Subscription

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        SubscriptionService susbscriptionService = Paymill.GetService<SubscriptionService>();

        Subscription subscription = new Subscription();
        subscription.Cancel_At_Period_End = false;
        subscription.Id = "sub_569df922b4506cd73030";

        Subscription updatedSubscription = susbscriptionService.UpdateSubscription(subscription);

        //Remove Subscription

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        SubscriptionService susbscriptionService = Paymill.GetService<SubscriptionService>();

        string subscriptionID = "sub_569df922b4506cd73030";
        bool reply = susbscriptionService.RemoveSubscription(subscriptionID);

        //List Subscriptions

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        SubscriptionService susbscriptionService = Paymill.GetService<SubscriptionService>();

        List<Subscription> lstSubscriptions = susbscriptionService.GetSubscriptions();

        //List Subscriptions with filter

        Paymill.ApiKey = "6bd9f94f4bd5f141da1aaea846d05ebd";
        Paymill.ApiUrl = "https://api.paymill.de/v2";
        SubscriptionService susbscriptionService = Paymill.GetService<SubscriptionService>();

        Filter filter = new Filter();
        filter.Add("count", 1); //OK
        filter.Add("offset", 2); //OK

        List<Subscription> lstSubscriptions = susbscriptionService.GetSubscriptions(filter);

        //Last edited Nov 28, 2012 at 7:23 PM by jcantos, version 12
                */
    }
}
