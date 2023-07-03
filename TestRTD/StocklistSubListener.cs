using com.lightstreamer.client;

namespace TestRTD
{
    internal class StocklistSubListener : SubscriptionListener
    {

        private IRtdLightstreamerListener listener;

        public StocklistSubListener(IRtdLightstreamerListener listener)
        {
            this.listener = listener;
        }

        public void onClearSnapshot(string itemName, int itemPos)
        {
            listener.OnItemUpdate(itemPos, itemName, null);
        }

        public void onCommandSecondLevelItemLostUpdates(int lostUpdates, string key)
        {
            // .
        }

        public void onCommandSecondLevelSubscriptionError(int code, string message, string key)
        {
            // .
        }

        public void onEndOfSnapshot(string itemName, int itemPos)
        {
            // .
        }

        public void onItemLostUpdates(string itemName, int itemPos, int lostUpdates)
        {
            listener.OnLostUpdate(itemPos, itemName, lostUpdates);
        }

        public void onItemUpdate(ItemUpdate itemUpdate)
        {
            listener.OnItemUpdate(itemUpdate.ItemPos, itemUpdate.ItemName, itemUpdate);
        }

        public void onListenEnd()
        {
            // .
        }

        public void onListenStart()
        {
            // .
        }

        public void onRealMaxFrequency(string frequency)
        {
            // .
        }

        public void onSubscription()
        {
            // .
        }

        public void onSubscriptionError(int code, string message)
        {
            // .
        }

        public void onUnsubscription()
        {
            // .
        }
    }
}