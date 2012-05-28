// Guids.cs
// MUST match guids.h
using System;

namespace Company.VSPackage1
{
    static class GuidList
    {
        public const string guidVSPackage1PkgString = "e35d2a97-22ec-4e92-ac32-13acaeed10e9";
        public const string guidVSPackage1CmdSetString = "a0e86bb1-dc6b-494b-a234-8e5bd1f5987d";

        public static readonly Guid guidVSPackage1CmdSet = new Guid(guidVSPackage1CmdSetString);
    };
}