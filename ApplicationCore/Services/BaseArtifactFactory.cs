﻿// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Text;
using ApplicationCore.Interfaces;
using ApplicationCore.Artifacts;
using ApplicationCore.Helpers;

namespace ApplicationCore.Services
{
    public abstract class BaseArtifactFactory<T> : IArtifactFactory<T> where T : BaseArtifact<T>
    {
        protected readonly ILogger _logger;
        protected readonly AppOptions _appOptions;

        public BaseArtifactFactory(
            ILogger logger,
            IOptions<AppOptions> appOptions)
        {
            Guard.Against.Null(logger, nameof(logger));
            Guard.Against.Null(appOptions, nameof(appOptions));

            _logger = logger;
            _appOptions = appOptions.Value;
        }
    }
}
