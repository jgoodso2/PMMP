﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace PMMP
{
    interface IDirector
    {
        Stream Construct(IBuilder builder, byte[] fileName, string projectGuid);
    }
}