/// <reference path="../typings/main.d.ts" />
import { expect } from 'chai';
import { mock } from 'sinon';
import { Graph } from '../src/KurveGraph';

describe('Graph', () => {

    describe('meAsync', () => {

        it('should get /me url', () => {
            var graph = new Graph({
                defaultAccessToken: 'access_token'
            });

            var mockedGet = mock(graph)
                .expects('get')
                .withArgs('https://graph.microsoft.com/v1.0/me/')

            graph.meAsync().then(() => {
                mockedGet.verify();
            });
        });
    });
});
