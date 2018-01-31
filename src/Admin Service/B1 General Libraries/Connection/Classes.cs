using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.Connection
{
    public class ConfigurationInformation
    {
        private string _ConnectionId = "";
        private string _ConnectionName = "";
        private string _Server = "";
        private string _LicenseServer = "";
        private SAPbobsCOM.BoDataServerTypes _DBServerType;
        private string _CompanyDB = "";
        private string _DBUserName = "";
        private string _DBUserPassword = "";
        private string _B1UserName = "";
        private string _B1Password = "";
        private string _Instance = "";
        private string _DefaultSchema = "";
        private string _UserName = "";
        private string _Password = "";
        private string[] _Type = new string[] { };

        public string ConnectionId
        {
            get { return _ConnectionId; }
            set
            {

                _ConnectionId = value;
            }
        }

        public string[] Type
        {
            get { return _Type; }
            set
            {

                _Type = value;
            }
        }
        public string ConnectionName { get { return _ConnectionName; } set { _ConnectionName = value; } }
        public string Server
        {
            get { return _Server; }
            set
            {
                if (Settings._Main.isEncrypted)
                {
                    //_Server = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    _Server = value;
                }

            }
        }
        public string LicenseServer
        {
            get { return _LicenseServer; }
            set
            {
                if (Settings._Main.isEncrypted)
                {
                    //_LicenseServer = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    _LicenseServer = value;
                }

            }
        }
        public SAPbobsCOM.BoDataServerTypes B1DBServerType { get { return _DBServerType; } }
        public string DBServerType
        {
            set
            {


                string strValue = "";


                if (Settings._Main.isEncrypted)
                {
                    //strValue = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    strValue = value;
                }

                _DBServerType = (SAPbobsCOM.BoDataServerTypes)Enum.Parse(typeof(SAPbobsCOM.BoDataServerTypes), strValue);
            }
        }
        public string CompanyDB
        {
            get { return _CompanyDB; }
            set
            {
                if (Settings._Main.isEncrypted)
                {
                    //_CompanyDB = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    _CompanyDB = value;
                }
            }
        }
        public string DBUserName
        {
            get { return _DBUserName; }
            set
            {
                if (Settings._Main.isEncrypted)
                {
                    //_DBUserName = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    _DBUserName = value;
                }
            }
        }
        public string DBUserPassword
        {
            get { return _DBUserPassword; }
            set
            {
                if (Settings._Main.isEncrypted)
                {
                    //_DBUserPassword = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    _DBUserPassword = value;
                }
            }
        }
        public string B1UserName
        {
            get { return _B1UserName; }
            set
            {
                if (Settings._Main.isEncrypted)
                {
                    //_B1UserName = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    _B1UserName = value;
                }
            }
        }
        public string B1Password
        {
            get { return _B1Password; }
            set
            {
                if (Settings._Main.isEncrypted)
                {
                    //_B1Password = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    _B1Password = value;
                }
            }
        }
        public string Instance
        {
            get { return _Instance; }
            set
            {
                if (Settings._Main.isEncrypted)
                {
                    //_Instance = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    _Instance = value;
                }
            }
        }
        public string DefaultSchema
        {
            get { return _DefaultSchema; }
            set
            {
                if (Settings._Main.isEncrypted)
                {
                    //_DefaultSchema = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    _DefaultSchema = value;
                }
            }
        }
        public string UserName
        {
            get { return _UserName; }
            set
            {
                if (Settings._Main.isEncrypted)
                {
                    //_UserName = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    _UserName = value;
                }
            }
        }
        public string Password
        {
            get { return _Password; }
            set
            {
                if (Settings._Main.isEncrypted)
                {
                    //_Password = SuSo.Encryption.Decrypt(value);
                }
                else
                {
                    _Password = value;
                }
            }
        }

        

    }
}
