namespace EPiServer.ConnectForSharePoint.Messaging
{
    public interface IMessageEmitterFactory
    {
        /// <summary>
        /// Gets the current number of items in the queue
        /// </summary>
        int QueueCount { get; }

        /// <summary>
        /// Creates an IMessageEmmiter which can then be used to create messages.
        /// </summary>
        /// <returns></returns>
        IMessageEmitter Create();

        /// <summary>
        /// Initializes the message message hander
        /// </summary>
        /// <param name="registry"></param>
        void Initialize(MessageHandlerRegistry registry);

        /// <summary>
        /// Clears all items in the queue.
        /// </summary>
        void Clear();
    }
}
