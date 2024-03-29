﻿using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomScript {
    public class CustomScriptItemMenuCommand : OleMenuCommand {
        private Predicate<int> matches;
        public string Path { get; set; }
        public CustomScriptItemMenuCommand(
            CommandID rootId,
            Predicate<int> matches,
            EventHandler invokeHandler,
            EventHandler beforeQueryStatusHandler) :
            base(invokeHandler, null, beforeQueryStatusHandler, rootId) {
            if (matches == null) {
                throw new ArgumentNullException("matches");
            }

            this.matches = matches;
        }

        public override bool DynamicItemMatch(int cmdId) {
            // Call the supplied predicate to test whether the given cmdId is a match.
            // If it is, store the command id in MatchedCommandid
            // for use by any BeforeQueryStatus handlers, and then return that it is a match.
            // Otherwise clear any previously stored matched cmdId and return that it is not a match.
            if (this.matches(cmdId)) {
                this.MatchedCommandId = cmdId;
                return true;
            }

            this.MatchedCommandId = 0;
            return false;
        }
    }
}
