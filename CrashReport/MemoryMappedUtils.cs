using System;
using System.Collections.Generic;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace CommonTabblesAndExplorerExtensions
{
    public static class MemoryMappedUtils
    {


        
        public const int defaultSizeInBytes = 5000000;

        public static byte[] getByteArr(MemoryMappedViewAccessor accessor)
        {
            byte[] buffer = new byte[accessor.Capacity];
            accessor.ReadArray<byte>(0, buffer, 0, buffer.Length);

            return buffer;
        }
        private static XDocument openXdocOrCreate(MemoryMappedViewStream fs)
        {
            XDocument xdoc;
            try
            {
                xdoc = XDocument.Load(fs);
            }
            catch (XmlException)
            {
                //Log("warning: xml file was incorrect, recreated. Messages could have been lost");
                var root = new XElement("root");
                xdoc = new XDocument(root);
            }
            return xdoc;
        }


        public static MemoryMappedFile getStringFromMmf(string mmfName, out string ret, int size = defaultSizeInBytes)
        {
            MemoryMappedFile mmf = creaOApriMmfFacendomiSapereSeLHaiCreato(mmfName, size);



            var accessor = mmf.CreateViewStream();

            try
            {


                accessor.Seek(0, SeekOrigin.Begin);

                var sr = new BinaryReader(accessor);

                var len = sr.ReadInt32();
                if (len == 0)
                {
                    ret = null;
                    return mmf;
                }


                byte[] buffer2 = new byte[len];
                accessor.Read(buffer2, 0, len);

                var enc = new UTF8Encoding();
                string str = enc.GetString(buffer2);


                ret = str;
                return mmf;
            }

            finally
            {
                accessor.Dispose();
            }

            //}
            //finally
            //{
            //    mmf.Dispose();
            //}



        }

        /// <summary>
        /// crea o apre un mmf. ci scrive dentro. torna l'mmf stesso perché il chiamante deve tenere un riferimento, altrimenti viene garbage collected.
        /// 
        /// l'mmf è a lunghezza fissa, quindi i dati potrebbero non entrare, nel qual caso torna errore.
        /// </summary>
        /// <param name="mmfName"></param>
        /// <param name="strToWrite"></param>
        /// <param name="size">se passi meno di 4096, crea comunque 4096</param>
        /// <returns>il memory mapped file creato. il chiamante deve tenere un riferimento, altrimenti viene garbage collected</returns>
        public static MemoryMappedFile writeStringToMmf2(string mmfName, string strToWrite,  out bool iDatiSonoEntratiNelFile, int size = defaultSizeInBytes)
        {

            var enc = new UTF8Encoding();
            byte[] buffer = enc.GetBytes(strToWrite);

            //var size = buffer.Length + 8000;


            MemoryMappedFile mmf = creaOApriMmfFacendomiSapereSeLHaiCreato(mmfName, size);


            var accessor = mmf.CreateViewStream();

            try
            {


                var sw = new BinaryWriter(accessor);

                sw.Write(buffer.Length); // prima scrivo la lunghezza dalla stringa in byte

                try
                {

                    sw.Write(buffer); // poi la stringa stessa
                    sw.Flush();
                }
                catch (NotSupportedException)
                {
                    iDatiSonoEntratiNelFile = false;
                    // la stringa è troppo lunga, non entra nel mmf. devo distruggerlo e ricrearlo più grande. AUTO GROW. non funziona.
                    //mmf.Dispose();
                    //accessor.SafeMemoryMappedViewHandle.Close();
                    //accessor.SafeMemoryMappedViewHandle.Dispose();
                    //accessor.Dispose();
                    //accessor.Close();
                    
                    
                }



                //// debug provo a leggere
                //accessor.Seek(0, SeekOrigin.Begin);

                //var sr = new BinaryReader(accessor);

                //var len = sr.ReadInt32();


                //byte[] buffer2 = new byte[len];
                //accessor.Read(buffer2, 0, len);
                //string strAppenaScritta = enc.GetString(buffer2);
                //var y = 4;

                iDatiSonoEntratiNelFile = true;
                return mmf;

            }
            finally
            {
                    accessor.Dispose();
            }

            //}
            //finally
            //{
            //    mmf.Dispose(); // no dispose! altrimenti distrugge il file e il contenuto va perso
            //}



        }

        /// <summary>
        /// la openOrCreate non mi permette di sapere se è stato aperto o creato, e quindi di debuggare.
        /// </summary>
        /// <param name="mmfName"></param>
        /// <param name="size">comunque non crea mai meno di 4096, anche se passi 3</param>
        /// <returns></returns>
        private static MemoryMappedFile creaOApriMmfFacendomiSapereSeLHaiCreato(string mmfName, int size)
        {
            MemoryMappedFile mmf;


            try
            {
                mmf = MemoryMappedFile.OpenExisting(mmfName);
            }
            catch (Exception)
            {
                var y = 4;  // breakpoint qui

                // dato che non esisteva, crealo.
                // non mi preoccupo delle race condition, perchè tanto questa funzione va chiamata in sezione critica con un mutex, quindi non
                // c'è pericolo che sia creato da qualcun altro tra le 2 istruzioni


                //var eccStr = eOpen.GetType().ToString() + " -- " + eOpen.Message + " " + (eOpen.StackTrace ?? "");
                //try
                //{

                mmf = MemoryMappedFile.CreateNew(mmfName, size);


                //}
                //catch(Exception eCreate)
                //{

                //    var eccCreateStr = eCreate.GetType().ToString() + " -- " + eCreate.Message + " " + (eCreate.StackTrace ?? "");


                //    throw new Exception("both open and create throw exception: open: " + eccStr + " --------- create: " + eccCreateStr);
                //}







                // debug vedi quanto l'ha creato grande . i lminimo è 4096

                //var acc = mmf.CreateViewAccessor();
                //var cap = acc.Capacity;
                //var z = 4;
            }

            return mmf;
        }
    }
}
