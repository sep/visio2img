using System;

namespace Visio2Img
{
    class Disposer : IDisposable
    {
        private readonly Action _onDispose;

        private Disposer(Action onDispose)
        {
            _onDispose = onDispose;
        }

        public static Disposer Create(Action onDispose)
        {
            return new Disposer(onDispose);
        }

        public void Dispose()
        {
            _onDispose();
        }
    }
}