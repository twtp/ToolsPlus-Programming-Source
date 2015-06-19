using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LivePriceComparer
{
        public partial class Form1
        {
            //make sure defaults coorelate to their setting checkboxes...
            private static bool FilterBlacklistedCompetitors = true;
            private static bool ShowPicture = true;
            private static bool ShowCompetitorName = true;
            private static bool ShowCompetitorFeedbackPercent = true;
            private static bool ShowCompetitorFeedbackScore = false;
            private static bool ShowProductTitle = true;
            private static bool ShowProductPricing = true;
            private static bool ShowShippingType = false;
            private static bool ShowExpedited = false;
            private static bool ShowOneDay = false;
            private static bool ShowShippingServices = false;
            private static bool ShowHandlingTime = false;
            private static bool ShowOnlyNewItems = true;
            ////////////////////////////////////////////////////////////

            //EBay Settings...
            private static bool EBayUseEBayAPI = true;
            private static bool EBayUseRawData = false;
            ////////////////////////////////////////////////////////////


            //Google Settings...
            private static bool GoogleSaveQueriedGIDs = true;
            private static bool GoogleJustUseFirstResultSet = true;
            private static bool GoogleUseAllResultSets = false;
            private static bool GoogleUseSavedGIDsOnly = false;
            private static bool GoogleUseSavedAndQueriedGIDs = false;            
            ////////////////////////////////////////////////////////////

            //Yahoo Settings...
            ////////////////////////////////////////////////////////////

            


            private void clearAllBrowsers()
            {
                ebayBrowser.Navigate("about:blank");
                debugBrowser.Navigate("about:blank");
                yahooBrowser.Navigate("about:blank");
                googleBrowser.Navigate("about:blank");                
            }

            private void filterBlacklistedChk_CheckedChanged(object sender, EventArgs e)
            {
                FilterBlacklistedCompetitors = filterBlacklistedChk.Checked;
            }

            private void showPictureChk_CheckedChanged(object sender, EventArgs e)
            {
                ShowPicture = showPictureChk.Checked;
            }

            private void showCompetitorNameChk_CheckedChanged(object sender, EventArgs e)
            {
                ShowCompetitorName = showCompetitorNameChk.Checked;
            }

            private void showFeedbackPercentChk_CheckedChanged(object sender, EventArgs e)
            {
                ShowCompetitorFeedbackPercent = showFeedbackPercentChk.Checked;
            }

            private void showFeedbackScoreChk_CheckedChanged(object sender, EventArgs e)
            {
                ShowCompetitorFeedbackScore = showFeedbackScoreChk.Checked;
            }

            private void showProductTitleChk_CheckedChanged(object sender, EventArgs e)
            {
                ShowProductTitle = showProductTitleChk.Checked;
            }

            private void showProductPricingChk_CheckedChanged(object sender, EventArgs e)
            {
                ShowProductPricing = showProductPricingChk.Checked;
            }

            private void showShippingTypeChk_CheckedChanged(object sender, EventArgs e)
            {
                ShowShippingType = showShippingTypeChk.Checked;
            }

            private void showExpeditedChk_CheckedChanged(object sender, EventArgs e)
            {
                ShowExpedited = showExpeditedChk.Checked;
            }

            private void showOneDayChk_CheckedChanged(object sender, EventArgs e)
            {
                ShowOneDay = showOneDayChk.Checked;
            }

            private void showServicesChk_CheckedChanged(object sender, EventArgs e)
            {
                ShowShippingServices = showServicesChk.Checked;
            }

            private void showHandlingChk_CheckedChanged(object sender, EventArgs e)
            {
                ShowHandlingTime = showHandlingChk.Checked;
            }
            
            private void setGoogleResultsSettings()
            {
                GoogleJustUseFirstResultSet = settingGoogleJustUseFirstResultSetRadio.Checked;
                GoogleUseAllResultSets = settingGoogleUseAllResultSetsRadio.Checked;
                GoogleUseSavedGIDsOnly = settingGoogleUseSavedGIDsOnlyRadio.Checked;
                GoogleUseSavedAndQueriedGIDs = settingGoogleUseSavedGIDsOrQueriedGIDsRadio.Checked;
            }
            
            private void settingGoogleJustUseFirstResultSetRadio_CheckedChanged(object sender, EventArgs e)
            {
                setGoogleResultsSettings();
            }

            private void settingGoogleUseAllResultSetsRadio_CheckedChanged(object sender, EventArgs e)
            {
                setGoogleResultsSettings();
            }

            private void settingGoogleUseSavedGIDsOnlyRadio_CheckedChanged(object sender, EventArgs e)
            {
                setGoogleResultsSettings();
            }

            private void settingGoogleUseSavedGIDsOrQueriedGIDsRadio_CheckedChanged(object sender, EventArgs e)
            {
                setGoogleResultsSettings();
            }

            private void settingGoogleSaveGIDsChk_CheckedChanged(object sender, EventArgs e)
            {
                GoogleSaveQueriedGIDs = settingGoogleSaveGIDsChk.Checked;
            }
            private void ebayUseEBayAPIRadio_CheckedChanged(object sender, EventArgs e)
            {
                EBayUseEBayAPI = true;
                EBayUseRawData = false;
            }
            private void ebayUseRawDataRadio_CheckedChanged(object sender, EventArgs e)
            {
                EBayUseRawData = true;
                EBayUseEBayAPI = false;
            }

        }
    

}
