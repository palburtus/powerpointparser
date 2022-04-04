using Microsoft.VisualStudio.TestTools.UnitTesting;
using Aaks.PowerPointParser.Html;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FluentAssertions;
using PowerPointParserTests.Html;

namespace Aaks.PowerPointParser.Html.Tests
{
    [TestClass]
    public class NestedHtmlListBuilderTests : BaseHtmlTests
    {
        [TestMethod]
        public void DoNotCloseListItemDueToNestingTest_NextIsNested_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.DoNotCloseListItemDueToNesting(
                BuildUnorderedListItem("two"),
                BuildUnorderedListItem("three", level:1));

            actual.Should().BeTrue();
        }

        [TestMethod]
        public void IsOnSameNestingLevel_NotNest_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.IsOnSameNestingLevel(
                BuildUnorderedListItem("one"),
                BuildUnorderedListItem("two"));

            actual.Should().BeTrue();
        }

        [TestMethod]
        public void IsOnSameNestingLevel_BothNested_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.IsOnSameNestingLevel(
                BuildUnorderedListItem("one", level:1),
                BuildUnorderedListItem("two", level: 1));

            actual.Should().BeTrue();
        }

        [TestMethod]
        public void IsOnSameNestingLevel_SecondNested_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.IsOnSameNestingLevel(
                BuildUnorderedListItem("one"),
                BuildUnorderedListItem("two", level: 1));

            actual.Should().BeFalse();
        }

        [TestMethod]
        public void IsOnSameNestingLevel_FirstNested_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.IsOnSameNestingLevel(
                BuildUnorderedListItem("one", level: 1),
                BuildUnorderedListItem("two"));

            actual.Should().BeFalse();
        }

        [TestMethod]
        public void IsOnSameNestingLevel_NextIsNull_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.IsOnSameNestingLevel(
                BuildUnorderedListItem("one"),null);

            actual.Should().BeTrue();
        }

        [TestMethod]
        public void ShouldChangeListTypes_NoTypeChanges_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.ShouldChangeListTypes(BuildUnorderedListItem("one"),
                BuildUnorderedListItem("two"),
                BuildUnorderedListItem("three"), "</ul>");

            actual.Should().BeFalse();
        }

        [TestMethod]
        public void ShouldChangeListTypes_CurrentUlPreviousOl_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.ShouldChangeListTypes(BuildUnorderedListItem("one"),
                BuildOrderListItem("two"),
                BuildUnorderedListItem("three"), "</ul>");

            actual.Should().BeTrue();
        }

        [TestMethod]
        public void ShouldChangeListTypes_NextNestedAndDifferent_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.ShouldChangeListTypes(
                BuildOrderListItem("one"),
                BuildUnorderedListItem("two", level: 1),
                BuildUnorderedListItem("three", level: 1), "</ul>");

            actual.Should().BeFalse();
        }

        [TestMethod]
        public void ShouldChangeListTypes_PreviousOneCurrentMatchesListType_ReturnsFalse()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.ShouldChangeListTypes(
                BuildOrderListItem("one", level:1),
                BuildUnorderedListItem("two"),
                BuildUnorderedListItem("three"), "</ul>");

            actual.Should().BeFalse();
        }

        [TestMethod]
        public void ShouldChangeListTypes_PreviousOneCurrentDoesNotMatchesListType_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.ShouldChangeListTypes(
                BuildOrderListItem("one", level: 1),
                BuildUnorderedListItem("two"),
                BuildUnorderedListItem("three"), "</ol>");

            actual.Should().BeTrue();
        }

        [TestMethod]
        public void ShouldChangeListTypes_CurrentAndPreviousLevelOneChangeType_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.ShouldChangeListTypes(
                BuildOrderListItem("one", level: 1),
                BuildUnorderedListItem("two"),
                BuildUnorderedListItem("three"), "</ol>");

            actual.Should().BeTrue();
        }

        [TestMethod]
        public void ShouldChangeListTypes_SwitchOnSameNestedLevel_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.ShouldChangeListTypes(
                BuildOrderListItem("one", level: 1),
                BuildUnorderedListItem("two", level: 1),
                BuildUnorderedListItem("three", level: 1), string.Empty);

            actual.Should().BeTrue();
        }

        [TestMethod]
        public void ShouldChangeListTypes_NestedLevelOneToTwo_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.ShouldChangeListTypes(
                BuildOrderListItem("one", level: 1),
                BuildUnorderedListItem("two", level: 2),
                BuildUnorderedListItem("three", level: 2), string.Empty);

            actual.Should().BeFalse();
        }

        [TestMethod]
        public void ShouldChangeListTypes_UnOrderedListDoesNotMatchOrderedListType_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.ShouldChangeListTypes(
                BuildOrderListItem("one", level: 1),
                BuildUnorderedListItem("two"),
                BuildUnorderedListItem("three"), "</ol>");

            actual.Should().BeTrue();
        }

        [TestMethod]
        public void ShouldChangeListTypes_SameListTypeDifferentBracket_ReturnsTrue()
        {
            INestedHtmlListBuilder builder = new NestedHtmlListBuilder();

            bool actual = builder.ShouldChangeListTypes(
                BuildUnorderedListItem("one", level: 1),
                BuildUnorderedListItem("two"),
                BuildUnorderedListItem("three"), "</ol>");

            actual.Should().BeTrue();
        }
    }
}