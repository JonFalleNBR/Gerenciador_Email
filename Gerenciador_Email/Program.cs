using System;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections.Generic;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using MailKit.Net.Smtp;

class Program
{
    static void Main()
    {
        // Configurações
        string email = "jonathan.leoesdejuda@gmail.com";
        string senhaApp = Environment.GetEnvironmentVariable("EMAIL_APP_PASSWORD");
        if (string.IsNullOrEmpty(senhaApp))
        {
            Console.WriteLine("Erro: Configure a variável de ambiente EMAIL_APP_PASSWORD.");
            Console.ReadLine();
            return;
        }

        int limiteDias = 30;
        string logFile = "email_cleanup_log.txt";

        // Lista de palavras-chave para remetentes
        var remetentes = new[] { "aliexpress", "claro", "udemy", "netflix","cruzeiro do sul", "amazon", "apple", "letterboxd", "ebay", "olx", "appbarber", "linkedin", "nubank", "spotify", "CEO", "bne","elem",
            "senai","vagas.com", "ciee", "catho", "youversion", "viotti", "infojobs", "discord", "quora", "trello", "ilikeyou", "disqus" , "marcos", "marina", 
        "abelssoft", "velox", "comic boom"};

        var dataLimite = DateTime.Now.AddDays(-limiteDias);
        int totalExcluidos = 0;

        var relatorio = new StringBuilder();
        relatorio.AppendLine($"📄 RELATÓRIO DE LIMPEZA DE EMAIL - {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n");

        // Iniciar contagem de tempo
        var startTime = DateTime.Now;

        try
        {
            using (var client = new ImapClient())
            {
                // Conectar ao servidor IMAP
                try
                {
                    client.Connect("imap.gmail.com", 993, true);
                    client.Authenticate(email, senhaApp);
                    Console.WriteLine("Conexão IMAP estabelecida com sucesso.");
                    LogToFile(logFile, "Conexão IMAP estabelecida com sucesso.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erro ao conectar ao IMAP: {ex.Message}");
                    LogToFile(logFile, $"Erro ao conectar/autenticar IMAP: {ex.Message}");
                    return;
                }

                // Somente a Caixa de Entrada (INBOX)
                var pastas = new[] { "INBOX", "[Gmail]/Spam", "[Gmail]/Lixeira" };

                foreach (var nomePasta in pastas)
                {
                    try
                    {
                        var pasta = client.GetFolder(nomePasta);
                        pasta.Open(FolderAccess.ReadWrite);
                        Console.WriteLine($"Processando pasta: {nomePasta}");
                        LogToFile(logFile, $"Processando pasta: {nomePasta}");

                        // Combinar filtros de remetentes com SearchQuery.Or
                        var remetenteQueries = remetentes.Select(r => SearchQuery.FromContains(r)).ToList();
                        SearchQuery remetenteQuery = null;
                        if (remetenteQueries.Any())
                        {
                            remetenteQuery = remetenteQueries[0]; // Primeira query
                            for (int i = 1; i < remetenteQueries.Count; i++)
                            {
                                remetenteQuery = SearchQuery.Or(remetenteQuery, remetenteQueries[i]);
                            }
                        }

                        // Buscar e-mails não lidos, com mais de 20 dias, de remetentes específicos
                        var uids = pasta.Search(SearchQuery.And(SearchQuery.And(SearchQuery.NotSeen, SearchQuery.DeliveredBefore(dataLimite)), remetenteQuery ?? SearchQuery.All));

                        // Recuperar metadados (remetente, assunto, data) em lote
                        var summaries = pasta.Fetch(uids, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId);

                        foreach (var summary in summaries)
                        {
                            Console.WriteLine($"Processando e-mail UID: {summary.UniqueId}"); // Log no console
                            LogToFile(logFile, $"Processando e-mail UID: {summary.UniqueId}");

                            Console.WriteLine($"Remetente: {summary.Envelope.From}"); // Log no console
                            LogToFile(logFile, $"Remetente: {summary.Envelope.From}");

                            // Verifica se o nome do remetente contém uma das palavras-chave
                            bool eRemetenteValido = summary.Envelope.From.Mailboxes
                                .Any(address => !string.IsNullOrEmpty(address.Name) &&
                                    remetentes.Any(r => address.Name.ToLower().Contains(r)));

                            if (eRemetenteValido)
                            {
                                Console.WriteLine($"Excluindo e-mail de: {summary.Envelope.From}"); // Log no console
                                LogToFile(logFile, $"Excluindo e-mail de: {summary.Envelope.From}");
                                pasta.AddFlags(summary.UniqueId, MessageFlags.Deleted, true);
                                totalExcluidos++;
                                relatorio.AppendLine($"🗑️ Excluído: {summary.Envelope.Subject} de {summary.Envelope.From} em {summary.Envelope.Date:yyyy-MM-dd}");
                                LogToFile(logFile, $"Excluído e-mail: {summary.Envelope.Subject} de {summary.Envelope.From}");
                            }
                        }

                        // Remover e-mails marcados como deletados
                        if (totalExcluidos > 0)
                        {
                            pasta.Expunge();
                            Console.WriteLine($"Expunge executado na pasta {nomePasta}.");
                            LogToFile(logFile, $"Expunge executado na pasta {nomePasta}.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erro ao processar pasta {nomePasta}: {ex.Message}");
                        LogToFile(logFile, $"Erro ao processar pasta {nomePasta}: {ex.Message}");
                        relatorio.AppendLine($"⚠️ Erro ao processar pasta {nomePasta}: {ex.Message}");
                    }
                }

                client.Disconnect(true);
                Console.WriteLine("Conexão IMAP encerrada.");
                LogToFile(logFile, "Conexão IMAP encerrada.");
            }

            // Gerar resumo do relatório
            if (totalExcluidos > 0)
            {
                relatorio.AppendLine($"\n🗑️ Total de {totalExcluidos} e-mails apagados.");
            }
            else
            {
                relatorio.AppendLine("\nNenhum e-mail apagado nesta execução.");
            }

            // Enviar e-mail com o relatório
            try
            {
                var msg = new MimeMessage();
                msg.From.Add(new MailboxAddress("", email));
                msg.To.Add(new MailboxAddress("", email));
                msg.Subject = "Relatório de Limpeza de E-mails";

                msg.Body = new TextPart("plain")
                {
                    Text = relatorio.ToString()
                };

                using (var smtp = new SmtpClient())
                {
                    smtp.Connect("smtp.gmail.com", 587, MailKit.Security.SecureSocketOptions.StartTls);
                    smtp.Authenticate(email, senhaApp);
                    smtp.Send(msg);
                    smtp.Disconnect(true);
                    Console.WriteLine("✅ Relatório enviado por e-mail.");
                    LogToFile(logFile, "Relatório enviado por e-mail com sucesso.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao enviar e-mail: {ex.Message}");
                LogToFile(logFile, $"Erro ao enviar e-mail: {ex.Message}");
            }

            // Log do tempo total
            var tempoTotal = (DateTime.Now - startTime).TotalSeconds;
            Console.WriteLine($"⏱️ Tempo total: {tempoTotal:F2} segundos");
            LogToFile(logFile, $"Tempo total: {tempoTotal:F2} segundos");

            Console.WriteLine(relatorio.ToString());
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro geral: {ex.Message}");
            LogToFile(logFile, $"Erro geral: {ex.Message}");
        }

        //Console.WriteLine("\nPressione ENTER para sair...");
        //Console.ReadLine();
    }

    // Função para registrar logs em arquivo
    static void LogToFile(string filePath, string message)
    {
        try
        {
            File.AppendAllText(filePath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao gravar log: {ex.Message}");
        }
    }
}