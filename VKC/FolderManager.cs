using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace BranchIntegrator
{
    class FolderManager
    {

        public void MovetoBackup()
        {
            try
            {
                General g = new General();
                DirectoryInfo diSource = new DirectoryInfo(g.path);


                string dt = DateTime.Now.ToString("ddMMyy-hhmmss");

                foreach (FileInfo fi in diSource.GetFiles())
                {

                    fi.MoveTo(g.pathBackup + dt + fi.Name);


                }
            }

            catch
            {
                MessageBox.Show ("Error While Moving File to Backup folder");
               
            }



        }
    }
}
