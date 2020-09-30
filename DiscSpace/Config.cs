using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DiscSpace
{
    public class Config
    {
        public List<string> NotifyTo { get; set; }

        public string NotifyIfFreeSpaceLessThan { get; set; }

        public int DelayBetweenTwoNotificationsInHours { get; set; }

        public DateTime LastNotificationSentTime { get; set; }
    }
}
