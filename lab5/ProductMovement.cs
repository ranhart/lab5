﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab5
{
    internal class ProductMovement
    {
        public int OperationID { get; set; }
        public DateTime Date { get; set; }
        public string ShopID { get; set; }
        public int Article { get; set; }
        public string OperationType { get; set; }
        public int ItemsQuantity { get; set; }
        public string Card { get; set; }

        public ProductMovement(int operationID, DateTime date, string shopID, int article, string operationType, int itemsQuantity, string card)
        {
            OperationID = operationID;
            Date = date;
            ShopID = shopID;
            Article = article;
            OperationType = operationType;
            ItemsQuantity = itemsQuantity;
            Card = card;
        }
        public ProductMovement(int operationID, DateTime date, string shopID, int article, string operationType, int itemsQuantity, string card)
        {
            OperationID = operationID;
            Date = date;
            ShopID = shopID;
            Article = article;
            OperationType = operationType;
            ItemsQuantity = itemsQuantity;
            Card = card;
        }

        public override string ToString()
        {
            return $"ID операции: {OperationID}, дата: {Date}, ID магазина: {ShopID}, артикул: {Article}, тип операции: {OperationType}, кол-во упаковок: {ItemsQuantity}, наличие карты клиента: {Card}"; 
        }
    }
}
