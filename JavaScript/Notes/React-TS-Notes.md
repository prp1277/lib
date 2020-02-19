# Getting Started with React and TypeScript

## The Server

`App.tsx` - Main landing page
`npm run start` to start the server

# Testing with Jest

Jest is the package that controls the testing
`npm run test` starts the testing server which can be run alongside `npm run start`. This allows you to preview changes and test them simultaneously.

## First, npm install Enzyme

First, we have to install [Enzyme](http://airbnb.io/enzyme/), which makes it easier to write tests for how components will behave. Enzyme builds on [jsdom](//TypeScript-React-Starter/tjs-testing/node_modules/jsdom) and makes it easier to make certain queries about our components.

We'll install it as a development-time dependency:
`npm install -D enzyme @types/enzyme enzyme-adapter-react-16 @types/enzyme-adapter-react-16 react-test-renderer`

- Enzyme package refers to the package containing the javascript code that actually gets to run
- @ types/enzyme is a package that contains declaration files `.d.ts` so that TypeScript can understand how to use Enzyme
- [Click here](https://www.typescriptlang.org/docs/handbook/declaration-files/consumption.html) for more information about the packages
- Enzyme-adapter-react-16 and react-test-renderer are enzyme dependencies

## Setting Up Our Tests

Create `src/setupTests.ts` - this will be automatically loaded when running tests

- [Here's](https://github.com/Microsoft/TypeScript-React-Starter#writing-tests-with-jest) the GitHub Readme Link

# The Production Build

A Minimized version of the app that creates an optimized JS and CSS build in `./build/static/js` and `./build/static/css`

# Classes

Classes are most useful when our component instances have some state or need to handle lifecycle hooks

# Adding State Management

React doesn't prescribe a specific way of synching data throughout the app. Data flows down through its children through the props specified on each element. Because of that, people generally use [Redux](http://redux.js.org/) or [MobX](https://mobx.js.org/).

## MobX

Written in TypeScript, MobX relies on functional reactive patterns where the state is wrapped through observables and passed through as props. Keeping the state fully synchronized for any for any observers is done by making the state observable.

## Redux

Synchronizes data through a central, immutable store of data and updates to that data will trigger re-renders. State gets updated in "an immutable fashion" by sending explicit messages, which must be handled by functions called reducers. This makes it easier to see how an action will affect the state of the program.

In this example, we'll use Redux. It's documentation can be found [here](http://redux.js.org/).

# Using Redux

If the state of the app doesn't change, using Redux doesn't make sense. Therefore, we'll need a source of actions that trigger the changes - ie a timer, button, etc.

First, install it - `npm install -S redux react-redux @types/react-redux`.

## Defining The App's State

Create a new directory and add an index.tsx file to it - `src/types/index.tsx`.

We'll use this to define the shape of the state that Redux will store. It will contain definitions for types that we may use throughout the program. For now, we're going to define the languageName and enthusiasmLevel.

Why we intentionally make our `state` slightly different than our `props` will make more sense when we move on to containers.

## Then, Add Actions

Create a new file and directory - `src/constants/index.tsx`. This is going to have some messages that our app can respond to.

The `const / type` pattern used in `constants/index` allows us to use TypeScripts string literal types in a way that they are easily accessible and refactorable. Next, well create a set of actions and functions that can create these actions.

This will also require a new directory - `src/actions/index.tsx`. Here, we describe the actions, add some logic (`EnthusiasmAction`) and manufacture the actions via functions. For more information see [redux-actions](https://www.npmjs.com/package/redux-actions).

## Adding a Reducer

Reducers are functions that generate changes by creating modified copies of the app's state, but have no side effects. These are called [pure functions](https://en.wikipedia.org/wiki/Pure_function).

Using the object spread `(... state)`, we create a shallow copy of our state while replacing `enthusiasmLevel`. Make sure that the `enthusiasmLevel` property comes last to avoid it getting overwritten by our old property's state.

It is important to test reducers to make sure they are producing the right state. For more information, see [Jest's toEqual](https://facebook.github.io/jest/docs/expect.html#toequalvalue) Method.

# Containers

Components for the most part are data agnostic used more for display. Containers wrap components and feed them any data necessary to display and modify state.

> [Presentational and Container Components](https://medium.com/@dan_abramov/smart-and-dumb-components-7ca2f9a7c7d0)

First, we have to update `src/components/Hello.tsx` by adding two optional callback properties to `Props`, then we'll bind those callbacks to two new buttons that we'll add to the component. These are the `onIncrement` and `onDecrement` properties.

After those have been added, we create `src/containers/Hello.tsx` to wrap the component. In the import statement, the two key pieces are the original `Hello` component as well as the `connect` function.

The `connect` function will take our `Hello` component and turn it into a container using two functions:
`mapStateToProps` - massages data from the current store to part of the shape the component needs
`mapDispatchToProps` - creates callback props to pump actions to our store using a given `dispatch` function

## Reviewing Our State

Our `state` consists of two properties: `languageName` and `enthusiasmLevel`, but our Hello component is expecting a `name` and an `enthusiasmLevel`. We will use `mapStateToProps` to retrieve the data from the store and adjust it, if necessary.

# Creating a Store

To bring everything full-circle, we will need to create a store with an initial state, then set it up with all the reducers.

`store` is our central store for the app's global state, which we'll add to `src/index.tsx`. Then, we'll swap the `component/Hello` with `containers/Hello`, using react-redux's `provider` to wire up the props with the container.

This is done by importing each into the container and passing our `store` through the `provider`s attributes.
