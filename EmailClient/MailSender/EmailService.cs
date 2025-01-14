﻿using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.IO;

namespace EmailCalendarsClient.MailSender
{
    public class EmailService
    {
        List<Attachment> MessageAttachmentsCollectionPage = new List<Attachment>();

        public Message CreateStandardEmail(string recipient, string header, string body)
        {
            var message = new Message
            {
                Subject = header,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = body
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipient
                        }
                    }
                },
                Attachments = MessageAttachmentsCollectionPage
            };

            return message;
        }

        public Message CreateHtmlEmail(string recipient, string header, string body)
        {
            var message = new Message
            {
                Subject = header,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = body
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipient
                        }
                    }
                },
                Attachments = MessageAttachmentsCollectionPage
            };

            return message;
        }

        public void AddAttachment(byte[] rawData, string filePath)
        {
            MessageAttachmentsCollectionPage.Add(new FileAttachment
            {
                Name = Path.GetFileName(filePath),
                ContentBytes = EncodeTobase64Bytes(rawData)
            });
        }

        public void ClearAttachments()
        {
            MessageAttachmentsCollectionPage.Clear();
        }

        static public byte[] EncodeTobase64Bytes(byte[] rawData)
        {
            string base64String = System.Convert.ToBase64String(rawData);
            var returnValue = Convert.FromBase64String(base64String);
            return returnValue;
        }
    }
}
